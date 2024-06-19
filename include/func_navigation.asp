<%
'******************************************************************************
'概　要：ヘッダー
'引　数：HeadType	0【トップ】1【求職者】2【企業】3【共用】4【代理店】
'作成者：Lis Niina
'作成日：2008/02/07
'備　考：
'使用元：
'******************************************************************************
Function NaviHeader(HeadType)
	Dim sHeadcmt
	Dim sLinkurl
	Dim sLinkalt
	Dim sLinktext

	Dim sContents
	Dim flgQE,oRS,sSQL,sError

	Dim cnt
	Dim iAll		'求職者数
	Dim iOrderCnt	'掲載中求人数
	Dim iCompanyCnt	'掲載中企業数


	iAll = 0
	iOrderCnt = 0
	iCompanyCnt = 0

	sSQL = ""
	sSQL = sSQL & "/* しごとナビ ＴＯＰページ用の出力データ取得 */"
	sSQL = sSQL & "EXEC up_DtlTopStatus;"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		iAll = oRS.Collect("StaffCnt")
		iOrderCnt = oRS.Collect("OrderCnt")
		iCompanyCnt = oRS.Collect("CompanyCnt")
	End If
	Call RSClose(oRS)





	If HeadType = 0 Then 'トップ
		sLinkurl = "/company/index.asp"
		sLinkalt = "求人広告しごとナビ"
		sLinktext = "採用担当（求人企業）様はこちら"
	ElseIf HeadType = 1 Then '求職者
		sLinkurl = "/company/index.asp"
		sLinkalt = "求人広告しごとナビ"
		sLinktext = "採用担当（求人企業）様はこちら"
	ElseIf HeadType = 2 Or HeadType = 4 Then '企業
		sLinkurl = "/"
		sLinkalt = "転職・求人サイトしごとナビ"
		sLinktext = "お仕事をお探しの方はこちら"
	ElseIf HeadType = 3 Then '共用
		sLinkurl = "/company/index.asp"
		sLinkalt = "求人広告しごとナビ"
		sLinktext = "採用担当（求人企業）様はこちら"
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
	<a id="smartLogo" href="/search/"><img src="/img/smart/smartLogo.png" alt="しごとナビ" border="0"></a>
	<a href="/order/order_search_detail.asp" id="smartSearch">探す</a>
    <div id="smartButton">
    </div>


</div><!--smartMenu-->
<div id="smartPhoneNavi" style="display:none;">
<% If G_USERTYPE = "staff" Then %>
	<h3>Myメニュー</h3>
    <ul>
    	<li class="topNavi"><a href="/staff/s_login.asp">Myページ</a></li>
		<li class="topNavi"><a href="/staff/person_detail.asp">プロフィール管理</a></li>
	<h3>転職サポート</h3>
        <li class="topNavi"><a href="/staff/my_footprint.asp">閲覧履歴</a></li>
        <li class="topNavi"><a href="/staff/watchlist.asp">お気に入りリスト</a></li>
		<li class="topNavi"><a href="/staff/edit_list.asp">応募一覧</a></li>
        <li class="topNavi"><a href="/staff/mailhistory_person.asp">メール管理</a></li>
        <li class="topNavi"><a href="/staff/schedule/">スケジュール管理</a></li>
        
	<h3>バリューオファー</h3>
        <li class="topNavi"><a href="/staff/step2a.asp">希望条件入力</a></li>
        <li class="topNavi"><a href="/staff/step2a.asp">企業からの質問</a></li>
	<h3>履歴書・職務経歴書の作成</h3>
		<li class="topNavi"><a href="/staff/resume_print.asp">履歴書・職務経歴書印刷</a></li>
		<li class="topNavi"><a href="/staff/resume_picture.asp">履歴書用写真登録</a></li>
	<h3>各種設定</h3>
        <%'<li class="topNavi"><a href="/staff/searchordercondition/">検索条件管理</a></li>%>
		<li class="topNavi"><a href="/staff/notification_mail_service.asp">スケジュール通知</a></li>
		<li class="topNavi"><a href="/staff/changepassword.asp">パスワードの変更</a></li>
		<li class="topNavi"><a href="/suspension/questionnarie.asp">休止・退会</a></li>


        <% 'なにこのif文
        If G_SSLFLAG = False Then %>
        <li class="topNavi"><a href="/logout.asp">ログアウト</a></li>
        <% ELSE %>
        <li class="topNavi"><a href="/logout.asp">ログアウト</a></li>
        <% END IF %>	
    	
    </ul>
    
<% End If %>
	<h3>メインメニュー</h3>
	<nav>
    	<ul>
            <li class="topNavi"><a href="/search/">しごとを探す</a></li>
            <li class="topNavi"><a href="/koryu/">交流</a></li>
            <li class="topNavi"><a href="/manabu/">学ぶ</a></li>
            <li class="topNavi"><a href="/link/">リンク</a></li>
        </ul>
    </nav>
    <h3>初めての方へ</h3>
    <nav>
    	<ul>
            <li class="topNavi"><a href="/tab/index1.asp">初めての方へ</a></li>
            <li class="topNavi"><a href="/valueoffer/">転職の新スタイル「バリューオファー」</a></li>
            <!--<li class="topNavi"><a href="/valueoffer/persona.asp">「バリューオファー物語～佐藤優一 編～」</a></li>-->
            <%'<li class="topNavi"><a href="/neo/howabout/">転職の新スタイル「エージェントNEO」</a></li>%>
        </ul>
    </nav>
    
    <br clear="both">
    <div id="smartNaviClose">
    	×CLOSE
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


<%  '2015/08/27 スマホかつログイン済みの時ヘッダーの一部を非表示
    '→色々なスクリプトの動作に絡むため中止
    'If chkSmartPhone(G_USERAGENT) = True and G_USERTYPE = "staff" Then
    'Response.Write "<div style=""display:none;"">" else %>
<div id="header_waku">
    <%' end if %>
	
    <div id="maku">    
    </div>
<header id="pagetop">


<div class="lt" id="top">
<a class="decnone" href="/"><img src="/img/top/logo.gif" alt="しごとナビ" border="0" align="left" style="margin-left:4px;"></a>
<br>
<p>はたらく人のソーシャルコミュニティー</p>
</div>

<div id="neoBanner">
    <!--<a href="/valueoffer/" id="toC">求職者様</a>-->
    <a href="/lis/lis_group.asp" id="toA">パートナー企業様</a>
 

</div>

<div class="rt">

<%

If G_USERTYPE = "" Then
	if GetForm("ordercode", 2) <> "" then
		if IsRE(Trim(Replace(Server.HTMLEncode(GetForm("ordercode", 2)), "'", "’")), "^J\d\d\d\d\d\d\d$", True) = True then
			response.write "<a href=""/staff/person_reg1.asp?ordercode=" & GetForm("ordercode", 2) & """target=""_self"" id=""reg_new"">会員登録</a>"
		else
			response.write "<a href=""/staff/person_reg1.asp"" target=""_self"" id=""reg_new"">会員登録</a>"
		end if
	else
		response.write "<a href=""/staff/person_reg1.asp"" target=""_self"" id=""reg_new"">会員登録</a>"
	end if


	    response.write "<a href=""/login_menu.asp"" target=""_self"" id=""login"">ログイン</a>"

ElseIf G_USERTYPE = "staff" Then
	response.write "<a href=""/staff/s_login.asp"" target=""_self"" id=""s_mypage"">My ﾍﾟｰｼﾞ</a>"

Else

End If
%>
<!--<a href="/staff/access.asp" class="stext"><img src="/img/top/head_icon.gif" height="10" alt="お問合せ" border="0">お問合せ</a>
<a href="/shigotonavi/sitemap.asp" class="stext">
<img src="/img/top/head_icon.gif" height="10" alt="サイトマップ" border="0">サイトマップ</a>
--></div>

<br clear="all">
<div style="position:absolute; right:25px; top:37px;">

<!-- #include file="../../caution.html" -->
</div>






<div id="number">
<%


	'<求人数、企業数、求職者数>

	Response.Write "求人<span class=""cnt"">" & iOrderCnt & "</span>件&nbsp;"
	Response.Write "企業<span class=""cnt"">" & iCompanyCnt & "</span>社&nbsp;"
	Response.Write "求職者<span class=""cnt"">" & iAll & "</span>人&nbsp;"
	Response.Write "（" & MonthName(Month(Now)) & Day(Now) & "日(" & Left(WeekdayName(Weekday(Now)),1) & ")" & "更新）</div>"
	'</求人数、企業数、求職者数>
	    
%>

<BR>
<div class="campaign" style="text-align:center;display:block;">
<strong style="color:#009900;font-size:250%;line-height:1.1em;text-align:center;display:inline;">おかげさまで『しごとナビ』登録者が１００万人突破しました。</strong>  
</div>



<div class="notice">
【しごとナビからのお知らせ】「新型コロナウイルス」感染拡大の予防対策として、お電話またはWEB面談(ZoomやSkype等)を利用しての非対面式での、<br> 転職の相談もご対応可能です。お気軽にお問い合わせください。

</div>


<%
	
If HeadType = 0 Then	
%>
<div id="img_map">

    <div id="comment_sagasu">
        <h4 class="center">「しごとを探す」とは</h4>
        求人情報検索や履歴書の自動作成<br>
        などができる、適職に就くための基本的な転職情報です。
    
    </div>
    
    <div id="comment_koryu">
        <h4 class="center">「交流」とは</h4>
        SNSで仕事や転職に関する情報や知識、アドバイスなどが得られるユーザー参加型の交流広場です。
        
    
    </div>
    
    <div id="comment_manabu">
        <h4 class="center">「学ぶ」とは</h4>
        スキルアップやビジネスコラム、<br>
        自己分析、転職ノウハウなどが学べるコーナーです。
    
    
    </div>
    
    <div id="comment_link">
        <h4 class="center">「リンク」とは</h4>
        当社が運営する関連サイトや、<br>
        関連情報が充実のお役立ちコーナーです。
    
    
    </div>
    
    <div id="comment_bes">
        <h4 class="center">会員登録（無料） をすると</h4>
        履歴書・経歴書の自動作成や自己分析ツールなど
        転職に役立つさまざまなコンテンツが使えます。
        企業からのスカウトメールを受け取ることもできます！
    
    
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
新バージョン
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
            <p>株式会社　NEO</p>
            <p>モバイル端末・ブロードバンドサービスの提案・販売</p>
            <p>東京都</p>
        </div>
        
        <div>
            <img src="https://www-b1.shigotonavi.co.jp/company/imgdsp.asp?companycode=C0018268&optionno=10">
            <p>株式会社　NEO</p>
            <p>モバイル端末・ブロードバンドサービスの提案・販売</p>
            <p>東京都</p>
        </div>
        
        <div>
            <img src="https://www-b1.shigotonavi.co.jp/company/imgdsp.asp?companycode=C0018268&optionno=10">
            <p>株式会社　NEO</p>
            <p>モバイル端末・ブロードバンドサービスの提案・販売</p>
            <p>東京都</p>
        </div>
        
        <div>
            <img src="https://www-b1.shigotonavi.co.jp/company/imgdsp.asp?companycode=C0018268&optionno=10">
            <p>株式会社　NEO</p>
            <p>モバイル端末・ブロードバンドサービスの提案・販売</p>
            <p>東京都</p>
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
        しごとナビの履歴書コンビニプリント.
    </p><br />
    
    <div style="text-align:center;margin:0 auto;background:#fff;">
    <!--<hr style="width:20px;border:1px solid #000;">-->
        <div style="padding:10px 0;background:#000;">
        <a href="/promotion/conpri_riyou.asp">  <img style="width:150px;border:0px solid #000;" src="/img/top/clogo_711.png"></a>
        <a href="/promotion/s_conpri_riyou.asp"><img style="width:150px;border:0px solid #000;" src="/img/top/clogo_familymart.png"></a>
        <a href="/promotion/s_conpri_riyou.asp"><img style="width:150px;border:0px solid #000;" src="/img/top/clogo_lawson.png"></a>
        </div>

    <p style="font-weight:bold;margin-top:12px;">サービスのご利用が可能なコンビニはこちら（五十音順）</p>
    </div>

</div>


    <!-- #INCLUDE VIRTUAL="/attention.html" -->
   <!-- <div id="mailReg">
    	<form>
        	<input type="text" value="ここにmail">
            <input type="button" value="登録" onClick="location.href='/staff/mailReg.asp'">
        </form>
    </div>-->
    <%
	End If
	



	response.write "<section id=""waku"">"


	If HeadType = 9 Then
		'<サイドメニュー無しver>
		'Response.Write "<div align=""left"" style=""width:100%;background-color:#ffffff;"">"
		'Response.Write "<div align=""left"" style=""width:990px;foat:left;"">"
		Response.Write "<div class=""moji912"" style=""padding:3px 0px 0px 3px;float:left;"">" & vbCrLf
		'</サイドメニュー無しver>
	Else
		'Response.Write "<div align=""left"" style=""width:100%;background-color:#ffffff;"">"
		'Response.Write "<div align=""left"" style=""width:990px;float:left;"">" 'ページ全体の幅（footer最下部で閉め
		Response.Write "<div class=""moji912"" id=""main"">" & vbCrLf 'メインコンテンツ幅（sidemenu最上部で閉め）
	End If
	
End Function

%><!-- #INCLUDE FILE="func/htmlNaviSideMenu.asp" --><%

'******************************************************************************
'概　要：フッター
'引　数：
'作成者：Lis Niina
'作成日：2008/02/07
'備　考：
'使用元：
'履　歴：2008/05/20 Lis林 しごとナビFC追加
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
	
	Response.Write "<p class=""m0"" style=""margin-top:15px;text-align:right;""><a href=""#pagetop"" class=""stext_bottom"">▲ページTOPへ</a></p>"
	Response.Write "<p class=""m0"" style=""margin-top:15px;text-align:right;""><a href=""#pagetop"" class=""stext"">▲ページTOPへ</a></p>"

	'<googleアドセンス>
	'2011/06/15～2011/06/21の期間はアドセンスをテスト的に停止する
	'If Date < "2011/06/15" Or Date >= "2011/06/22" Then
	'2011/07/08～ アドセンスを停止する
	If Date < "2011/07/08" Then
		Response.Write "<div style=""margin-bottom:10px;"">"
%><!-- #INCLUDE VIRTUAL="/include/ads/navifooter.asp" --><%
		Response.Write "</div>"
	End If
	'</googleアドセンス>


	'しごとナビモバイルの紹介（携帯のアドレス登録者のみ）
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
	<li class="ttl">「しごとナビ」について</li>
	<li><a href="<%= HTTP_CURRENTURL %> " class="topdecnone">しごとナビHOME</a></li>
    <li><a href="<%= HTTPS_CURRENTURL %>tab/index1.asp" class="topdecnone">はじめての方へ</a></li>
    <li><a href="<%= HTTPS_CURRENTURL %>search/" class="topdecnone">しごとを探すエリア</a></li>
    <li><a href="<%= HTTP_CURRENTURL %>koryu/" class="topdecnone">交流エリア</a></li>
    <li><a href="<%= HTTP_CURRENTURL %>manabu/" class="topdecnone">学ぶエリア</a></li>
    <li><a href="<%= HTTP_CURRENTURL %>link/" class="topdecnone">リンクエリア</a></li>
    <li><a href="<%= HTTPS_CURRENTURL %>support/" class="topdecnone">転職サポート</a></li>
    <li><a href="<%= HTTPS_CURRENTURL %>staff/ranking_index.asp" class="topdecnone">しごとナビランキング</a></li>
    <li><a href="<%= HTTPS_CURRENTURL %>/staff/s_aboutnavi.asp" class="topdecnone">ご利用ガイド</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>lis/lis.asp" class="topdecnone">運営会社について</a></li>
	<li><a href="<%= HTTPS_CURRENTURL %>recruit/" class="topdecnone">採用情報</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>shigotonavi/sitemap.asp" class="topdecnone">サイトマップ</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>privacy/privacymark.asp" class="topdecnone">Pマークについて</a></li>
	</ul>

	<ul>
	<li class="ttl">求職者様</li>
	<li><a href="<%= HTTP_CURRENTURL %>order/order_search_detail.asp" class="topdecnone">求人を探す</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>staff/s_resume.asp" class="topdecnone">履歴書の自動作成/フォーマットの<br>&nbspダウンロード</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>staff/s_resume_kakikata.asp" class="topdecnone">履歴書の書き方</a></li>
    	<li><a href="<%= HTTPS_CURRENTURL %>column/column_index.asp" class="topdecnone">転職・就職コラム</a></li>
    	<li><a href="<%= HTTPS_CURRENTURL %>type_map.asp" class="topdecnone">職種業種別マップ</a></li>
	<li><a href="<%= HTTPS_CURRENTURL %>s_contents/s_jikopr.asp" class="topdecnone">自己PRメーカー</a></li>
	<li><a href="<%= HTTPS_CURRENTURL %>s_contents/motive_index.asp" class="topdecnone">志望動機メーカー</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>staff/s_careersheet.asp" class="topdecnone">職務経歴書の自動作成/フォーマットの<br>&nbspダウンロード</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>staff/s_careersheet_kakikata_1.asp" class="topdecnone">職務経歴書の書き方</a></li>
	<li><a href="<%= HTTPS_CURRENTURL %>s_contents/s_mynavi.asp" class="topdecnone">適職診断「じぶんナビ」</a></li>
	<li><a href="<%= HTTPS_CURRENTURL %>s_contents/s_temporary.asp" class="topdecnone">人材派遣</a>｜<a href="<%= HTTPS_CURRENTURL %>s_contents/s_introduce.asp" class="topdecnone">人材紹介</a>｜<a href="<%= HTTPS_CURRENTURL %>s_contents/s_temptoperm.asp" class="topdecnone">紹介予定派遣</a></li>
	<li><a href="<%= HTTPS_CURRENTURL %>staff/access.asp" class="topdecnone">お問合せ</a></li>
	</ul>

	<ul>
	<li class="ttl">特集</li>
	<li><a href="<%= HTTP_CURRENTURL %>order/special/ad/0001/" class="topdecnone">SEの転職特集</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>order/special/tg/0004/" class="topdecnone">臨床検査技師の求人</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>order/special/tg/0005/" class="topdecnone">英語を活かして派遣で働く</a></li>
    <li><a href="<%= HTTP_CURRENTURL %>order/special/or/0001/" class="topdecnone">DTP・デザイナーの求人</a></li>
    <li><a href="<%= HTTP_CURRENTURL %>s_contents/license/1700101.asp" class="topdecnone">宅地建物取引主任者の求人</a></li>
    <li><a href="<%= HTTP_CURRENTURL %>order/special/tg/0006/index.asp" class="topdecnone">年収1000万円クラスの転職</a></li>
    <li><a href="<%= HTTP_CURRENTURL %>order/special/tk/0001/index.asp" class="topdecnone">建築・不動産業界特集</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>order/special/ng000001.asp" class="topdecnone">企業看護師の求人特集</a></li>
    <li><a href="<%= HTTP_CURRENTURL %>order/special/tokyo.asp" class="topdecnone">東京の転職・就職</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>order/special/sz/0001/" class="topdecnone">静岡の転職</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>order/special/ng/0002/" class="topdecnone">名古屋の転職</a></li>
    <li><a href="<%= HTTP_CURRENTURL %>order/special/oy/0001/" class="topdecnone">岡山の転職</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>order/special/hr/0001/" class="topdecnone">広島の転職</a></li>
	</ul>

	<ul style="margin-right:0px;">
	<li class="ttl">採用企業様</li>
    <li><a href="<%= HTTP_CURRENTURL %>neo/shoukai/" class="topdecnone">採用企業トップ</a></li>
    <li><a href="<%= HTTP_CURRENTURL %>company/" class="topdecnone">「エージェントNEO」について</a></li>
    <li><a href="<%= HTTP_CURRENTURL %>company/about.asp" class="topdecnone">しごとナビの特色</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>company/c_staffdata.asp" class="topdecnone">求職者集計データ</a></li>
	<li><a href="<%= HTTP_CURRENTURL %>company/c_dispatch.asp" class="topdecnone">人材派遣</a>｜<a href="<%= HTTP_CURRENTURL %>company/c_introduce.asp" class="topdecnone">人材紹介</a>｜<a href="<%= HTTP_CURRENTURL %>company/c_temptoperm.asp" class="topdecnone">紹介予定派遣</a></li>
    <!--<li><a href="<%= HTTPS_CURRENTURL %>neo/kokoku/advertisement.asp" class="topdecnone">求人広告のお申込み</a></li>-->
    <li><a href="<%= HTTPS_CURRENTURL %>neo/shoukai/index.asp" class="topdecnone">人材紹介のお申込み</a></li>
    <li><a href="<%= HTTPS_CURRENTURL %>company/research.asp" class="topdecnone">採用方法診断</a></li>
	<li><a href="<%= HTTPS_CURRENTURL %>company/access.asp" class="topdecnone">お問合せ</a></li>

    <li><a href="<%= HTTPS_CURRENTURL %>neo/TempRegist/TempRegistEdit_AD.asp" class="topdecnone">求人広告サービスについて</a></li>
	</ul>

	<br clear="all">

	<div style="text-align:center;">
	<a href="http://tekiseika.jp/job-offering/" target="_blank"><img src="/img/tekiseika_job-offering2.jpg" alt="求人者の皆さま" border="0" style="margin-top:3px;"></a>
	<a href="http://tekiseika.jp/job-applicant/" target="_blank"><img src="/img/tekiseika_job-applicant2.jpg" alt="求職者の皆さま" border="0" style="margin-top:3px;"></a>
	</div>
	
	<div style="text-align:center;">
	<a href="<%= HTTP_LIS_CURRENTURL %>" target="_blank"><img src="/img/footer/footer_lis_logo_1.gif" alt="転職サイト｢しごとナビ｣運営-リス株式会社-" border="0"></a>
	</div>
	
    <div id="smartFooter" style="display:none;">
    
    
    	<p>CopyRights(c)LIS co.,ltd.</p>
    </div>
</div>
	</footer>

<%    
    	'<スマートフォンユーザ向けのしごとナビモバイルへの誘導バナー表示>
'	If chkSmartPhone(G_USERAGENT) = True Then
'		'Response.Write "<a href=""" & HTTPS_NAVI_MOBILE & "?an=spbanner""><img src=""/img/banner/smartphone_banner.png"" alt=""スマートフォンの方はココをタッチ！しごとナビモバイル"" border=""0""></a>"
'        Response.Write "<div style=""padding:15px;line-height:2em;font-size:xx-large;"">"
'        Response.Write "<a href=""http://sp.shigotonavi.jp/"" border=""0""><img src=""/img/switch_btn_01.gif"" border=""0""></a>"
'        Response.Write "<img src=""/img/switch_btn_02.gif"" border=""0"">"
'        'Response.Write "PC | <a href=""http://sp.shigotonavi.jp/"">スマートフォン</a>"
'        Response.Write "</div>"
'
'	End If
	'</スマートフォンユーザ向けのしごとナビモバイルへの誘導バナー表示>
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

<% '2017/08/22 YSS用リマーケティングタグ追加 %>
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
<% '2017/08/22 YSS用リマーケティングタグ追加 %>

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
'概　要：右サイド
'引　数：
'作成者：Lis Niina
'作成日：2008/02/07
'備　考：
'使用元：
'履　歴：
' 08/05/20 Lis林 しごとナビFC追加
'******************************************************************************
Function NaviSidemenuRight()
	Dim oRSnsr,sSQLnsr,sErrornsr,flgQEnsr

	Response.Write "<div style=""width:200px;height:135px;background-image:url(/img/rightmenu/navicafe_banner_all.jpg);margin-bottom:5px;"">"
	Response.Write "<a href=""" & HTTP_CURRENTURL & "cafe/cafe_list.asp""><img src=""/img/rightmenu/navicafe_banner_top.jpg"" alt=""ナビカフェ"" border=""0"" style=""margin:0px;padding:0px;""></a>"
	Response.Write "<div style=""margin-top:0px;padding:14px 6px 0px 8px;font-size:10px;line-height:15px;"">"

	'** TOP 08/11/05 Lis林 ADD
	'現在掲載中＆TOP3のトピ
	sSQLnsr = "up_GetData_NC_Topic '','','','1','3'"
	flgQEnsr = QUERYEXE(dbconn, oRSnsr, sSQLnsr, sErrornsr)
	Do While GetRSState(oRSnsr) = True
		Response.Write "<a href=""" & HTTP_CURRENTURL & "cafe/cafe_detail.asp?t=" & oRSnsr.Collect("TopicID")
		Response.Write """>・"
		If Len(oRSnsr.Collect("Title")) > 14 Then
			Response.Write Left(oRSnsr.Collect("Title"),14) & "..."
		Else
			Response.Write oRSnsr.Collect("Title")
		End If
		Response.Write "</a><br>"
		oRSnsr.MoveNext
	Loop
	Call RSClose(sSQLnsr)
	'** BTM 08/11/05 Lis林 ADD

	Response.Write "</div>"
	Response.Write "</div>"

	If Session("usertype") = "staff" Then '求職者ログインしている場合	
		Response.Write "<ul>"
		Response.Write "<li class=""rightmenu_big"">サポート</li>"
		Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "staff/s_aboutnavi.asp"">ご利用ガイド</a></li>"
		Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "staff/qa.asp"">Ｑ＆Ａ</a></li>"
		Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "staff/s_searchexplanation.asp"">お仕事検索方法</a></li>"
		Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "staff/s_kiyaku.asp"">利用規約</a></li>"
		Response.Write "<li class=""rightmenu_end""><a href=""" & HTTPS_CURRENTURL & "staff/access.asp"">お問合せ(求職者専用)</a></li>"
		Response.Write "<li class=""rightmenu_bottom""></li>"
		Response.Write "</ul>"
	End If

	Response.Write "<ul>"
	Response.Write "<li class=""rightmenu_big"">ケータイでもしごとナビ</li>"
	Response.Write "<li style=""border-left:solid 1px #cccccc; border-right:solid 1px #cccccc;""><a href=""" & HTTP_CURRENTURL & "promotion/mobilepromotion.asp"" style=""display:block;text-align:center;""><img src=""/img/sidemenu/mobile_banner.jpg"" alt=""しごとナビモバイル"" border=""0""></a></li>"
	Response.Write "<li class=""rightmenu_bottom"" style=""clear:both;""></li>"
	Response.Write "</ul>"

	Response.Write "<ul>"
	Response.Write "<li class=""rightmenu_big"">ＣｏｎＰｒｉ（コンプリ）</li>"
	Response.Write "<li style=""height:51px; border-left:solid 1px #cccccc; border-right:solid 1px #cccccc; border-bottom:solid 1px #eeeeee;""><a href=""" & HTTP_CURRENTURL & "promotion/conpripromotion.asp"" style=""display:block;text-align:center;""><img src=""/img/rightmenu/conpri_banner1.jpg"" alt=""コンプリ"" border=""0""></a></li>"
	Response.Write "<li style=""border-left:solid 1px #cccccc; border-right:solid 1px #cccccc; padding:2px 3px; font-size:10px;"">パソコン、または携帯から作成した履歴書をコンビニで印刷できる画期的サービス！証明写真も取り込める！</li>"
	Response.Write "<li style=""border-left:solid 1px #cccccc; border-right:solid 1px #cccccc;""><a href=""" & HTTP_CURRENTURL & "promotion/conpripromotion.asp"" style=""display:block;text-align:center;""><img src=""/img/rightmenu/conpri_banner2.jpg"" alt=""詳しくはこちら"" border=""0""></a></li>"
	Response.Write "<li class=""rightmenu_bottom"" style=""clear:both;""></li>"
	Response.Write "</ul>"
%><!-- #include VIRTUAL="/include/ads/navirighttext.asp" --><%
	Response.Write "<ul>"
	Response.Write "<li class=""rightmenu_big"">コラム</li>"
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_mensetsu_index.asp"">面接対策</a></li>"
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "column/column_1.asp"">派遣社員<span class=""stext"">-成功の鍵はプロ意識</span></a></li>"
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_kyuuyomeisai.asp"">あなたの給与明細</a></li>"
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_ready.asp"">転職の心構え</a></li>"
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_proce.asp"">転職に必要な手続き</a></li>"
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_goukaku.asp"">合格率UPマニュアル</a></li>"
	Response.Write "<li class=""rightmenu_bottom""></li>"
	Response.Write "</ul>"
%><!-- #include VIRTUAL="/include/ads/navirightview.asp" --><%
	Response.Write "<br>"
	Response.Write "<div align=""center"" style=""width:100%;"">"
	Response.Write "<div class=""rightmenu_big"" style=""text-align:left;"">求職者情報</div>"
	Response.Write "<div style=""border-left:solid 1px #cccccc; border-right:solid 1px #cccccc; background-image:url(/img/sidemenu/jinzaidata_background.gif);"" align=""center"">"
	Response.Write "<table style=""width:155px; font-size:10px; text-align:left;"">"

	Dim rank(2)
	Dim rankcount(2)
	Dim idx
	idx = 0

	sSQLnsr = "SELECT top 3 Subitem,Number FROM Person_Statistics where item = '都道府県別' order by convert(int,Number) desc"
	flgQEnsr = QUERYEXE(dbconn, oRSnsr, sSQLnsr, sErrornsr)
	Do While GetRSState(oRSnsr) = True
		rank(idx) = Replace(Replace(Replace(oRSnsr.Collect("SubItem"),"都",""),"府",""),"県","")
		rankcount(idx) = oRSnsr.Collect("Number")
		idx = idx + 1
		oRSnsr.MoveNext
	Loop
	Call RSClose(oRSnsr)

	Response.Write "<tr>"
	Response.Write "<td>都道府県別</td>"
	Response.Write "<td>1位:" & rank(0) & "</td>"
	Response.Write "<td align=""right"">" & rankcount(0) & "名</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td></td>"
	Response.Write "<td>2位:" & rank(1) & "</td>"
	Response.Write "<td align=""right"">" & rankcount(1) & "名</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td></td>"
	Response.Write "<td>3位:" & rank(2) & "</td>"
	Response.Write "<td align=""right"">" & rankcount(2) & "名</td>"
	Response.Write "</tr>"

	idx = 0

	sSQLnsr = "SELECT top 3 item,subitem, Number FROM Person_Statistics where item = '10歳代' or item = '20歳代' or item = '30歳代' or item = '40歳代' or item = '50歳代' or item = '60歳以上' order by convert(int,Number) desc"
	flgQEnsr = QUERYEXE(dbconn, oRSnsr, sSQLnsr, sErrornsr)
	Do While GetRSState(oRSnsr) = True
		rank(idx) = Replace(oRSnsr.Collect("Item"),"歳","") & oRSnsr.Collect("SubItem")
		rankcount(idx) = oRSnsr.Collect("Number")
		idx = idx + 1
		oRSnsr.MoveNext
	Loop
	Call RSClose(oRSnsr)

	Response.Write "<tr>"
	Response.Write "<td>年齢別</td>"
	Response.Write "<td>1位:" & rank(0) & "</td>"
	Response.Write "<td align=""right"">" & rankcount(0) & "名</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td></td>"
	Response.Write "<td>2位:" & rank(1) & "</td>"
	Response.Write "<td align=""right"">" & rankcount(1) & "名</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td></td>"
	Response.Write "<td>3位" & rank(2) & "</td>"
	Response.Write "<td align=""right"">" & rankcount(2) & "名</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td colspan=""3"" align=""right""><a href=""" & HTTP_CURRENTURL & "company/c_staffdata.asp""><img src=""/img/sidemenu/kuwashiku_min.jpg"" alt=""詳しくはこちら"" border=""0""></a>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "</div>"
	Response.Write "<div class=""rightmenu_bottom"" style=""clear:both;""></div>"
	Response.Write "<br>"


End Function
%>
