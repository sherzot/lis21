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

	Response.Write "<a name=""pagetop""></a>" & vbCrLf

	sHeadcmt = getHeaderText(HeadType,G_URL)

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



	'<スマートフォンユーザ向けのしごとナビモバイルへの誘導バナー表示>
	If chkSmartPhone(G_USERAGENT) = True Then
		'Response.Write "<a href=""" & HTTPS_NAVI_MOBILE & "?an=spbanner""><img src=""/img/banner/smartphone_banner.png"" alt=""スマートフォンの方はココをタッチ！しごとナビモバイル"" border=""0""></a>"
        Response.Write "<div style=""padding:15px;line-height:2em;font-size:xx-large;"">"
        Response.Write "<a href=""http://sp.shigotonavi.jp/"" border=""0""><img src=""/img/switch_btn_01.gif"" border=""0""></a>"
        Response.Write "<img src=""/img/switch_btn_02.gif"" border=""0"">"
        'Response.Write "PC | <a href=""http://sp.shigotonavi.jp/"">スマートフォン</a>"
        Response.Write "</div>"

	End If
	'</スマートフォンユーザ向けのしごとナビモバイルへの誘導バナー表示>

%>
<div id="waku">
<header>

<div class="hblk1"></div>
<div class="lt">
<h1>職サイト「しごとナビ」。正社員・派遣の求人情報はもちろん、プロによる貴方に適した転職サポートをご提供しています！</h1>
</div>
<div class="rt">
<a href="/staff/access.asp" class="stext"><img src="/img/top/head_icon.gif" height="10" alt="お問合せ" border="0">お問合せ</a>
<a href="/shigotonavi/sitemap.asp" class="stext">
<img src="/img/top/head_icon.gif" height="10" alt="サイトマップ" border="0">サイトマップ</a>
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
		Response.Write "<a class=""decnone"" href=""/""><img src=""/img/top/shigotonavi_logo.gif"" alt=""しごとナビ"" border=""0"" align=""left"" style=""margin-left:4px;""></a>"
	End If
%>

</td>
	<!--/ヘッダー左：しごとナビロゴ-->

	<!--ヘッダー右-->
<td align="right" valign="bottom" style="font-size:11px;" class="topstatus">

	
    <%
	
	'<Googleのサイト内検索>
	If Request.ServerVariables("HTTPS") <> "on" Then
		Response.Write "<form action=""/search.asp"" id=""cse-search-box"" style=""margin-left:5px;padding:0px;display:inline;"">"
		Response.Write "<div style=""display:inline;"">"
		Response.Write "<img src=""/img/top/head_icon.gif"" alt="""" border=""0"" style=""vertical-align:millde;"">"
		Response.Write "<label>"
		Response.write "<span>サイト内検索&nbsp;&nbsp;</span>"
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
	'</Googleのサイト内検索>


	'<求人数、企業数、求職者数>
	Response.Write "<img src=""/img/top/countericon_order.gif"" alt=""求人数"" border=""0"" style=""margin:0px 2px;"">求人<span class=""cnt"">" & iOrderCnt & "</span>件&nbsp;"
	Response.Write "<img src=""/img/top/countericon_company.gif"" alt=""企業数"" border=""0"" style=""margin:0px 2px;"">企業<span class=""cnt"">" & iCompanyCnt & "</span>社&nbsp;"
	Response.Write "<img src=""/img/top/countericon_staff.gif"" alt=""求職者数"" border=""0"" style=""margin:0px 2px;"">求職者<span class=""cnt"">" & iAll & "</span>人&nbsp;"
	Response.Write "（" & MonthName(Month(Now)) & Day(Now) & "日(" & Left(WeekdayName(Weekday(Now)),1) & ")" & "更新）"
	'</求人数、企業数、求職者数>

	'採用担当者様
	'Response.Write "　<a href=""" & sLinkurl & """ style=""font-size:14px;""><img src=""/img/top/head_icon.gif"" alt=""" & sLinkalt & """ border=""0"" style=""vertical-align:middle;"">" & sLinktext & "</a>"
	'<!-- #INCLUDE FILE="../ad_banner_control/ad_banner.asp" -->
	Response.Write "</td>"
	'<ヘッダー右>

	Response.Write "</tr>"

	'<ヘッダー下部：背景緑のやつ>
'	Response.Write "<tr style=""background-image:url(/img/top/headtext_background.gif);"">"
'	Response.Write "<td colspan=""2"" align=""left"" style=""margin:0px;padding:0px;color:#ffffff; height:20px;border-top:solid 1px #ffffff; border-bottom:solid 1px #ffffff;"">"
'	Response.Write sHeadcmt
'	Response.Write "</td>"
'	Response.Write "</tr>"
	'</ヘッダー下部：背景緑のやつ>

	Response.Write "</table>"
	Response.Write htmlTabIndex(Request.ServerVariables("URL"),G_USERTYPE,sHeadcmt)
	Response.Write "</header>"

	If HeadType = 9 Then
		'<サイドメニュー無しver>
		Response.Write "<div align=""left"" style=""width:100%;background-color:#ffffff;"">"
		Response.Write "<div align=""left"" style=""width:990px;foat:left;"">"
		Response.Write "<div class=""moji912"" style=""padding:3px 0px 0px 3px;float:left;"">" & vbCrLf
		'</サイドメニュー無しver>
	Else
		Response.Write "<div align=""left"" style=""width:100%;background-color:#ffffff;"">"
		Response.Write "<div align=""left"" style=""width:990px;float:left;"">" 'ページ全体の幅（footer最下部で閉め
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
	Server.Execute("/include/mobilesiteinfo.asp")

	Response.Write "<div id=""footer"">"

	Response.Write "<ul>"
	Response.Write "<li class=""ttl"">転職サイト「しごとナビ」について</li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & """ class=""topdecnone"">転職サイト「しごとナビ」HOME</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "infomation/info.asp"" class=""topdecnone"">記事倉庫</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "lis/lis.asp"" class=""topdecnone"">運営会社・当社採用情報</a></li>"
'	Response.Write "<li><a href=""/staff/s_aboutnavi.asp"" class=""topdecnone"">ご利用ガイド</a></li>"
'	Response.Write "<li><a href=""/staff/qa.asp"" class=""topdecnone"">Ｑ＆Ａ</a></li>"
'	Response.Write "<li><a href=""/staff/s_kiyaku.asp"" class=""topdecnone"">利用規約</a></li>"
'	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "s_contents/s_books.asp"" class=""topdecnone"">転職に役立つ本</a></li>"
'	Response.Write "<li><a href=""/link.asp"" class=""topdecnone"">リンクポリシー</a></li>"
'	Response.Write "<li><a href=""/link_collection.asp"" class=""topdecnone"">お役立ち厳選リンク集</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "shigotonavi/sitemap.asp"" class=""topdecnone"">サイトマップ</a></li>"
'	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "staff/Ranking.asp"" class=""topdecnone"">求職者ランキング</a></li>"
	Response.Write "</ul>"

	Response.Write "<ul>"
	Response.Write "<li class=""ttl"">転職をお考えの求職者様</li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "order/order_search_detail.asp"" class=""topdecnone"">求人を探す</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "staff/s_resume.asp"" class=""topdecnone"">履歴書の自動作成ツール</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "staff/s_resume_kakikata.asp"" class=""topdecnone"">履歴書の書き方</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "s_contents/s_jikopr.asp"" class=""topdecnone"">自己ＰＲメーカー</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "s_contents/motive_index.asp"" class=""topdecnone"">志望動機メーカー</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "staff/s_careersheet.asp"" class=""topdecnone"">職務経歴書の自動作成ツール</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "staff/s_careersheet_kakikata_1.asp"" class=""topdecnone"">職務経歴書の書き方</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "s_contents/s_mynavi.asp"" class=""topdecnone"">適職診断「じぶんナビ」</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "s_contents/s_temporary.asp"" class=""topdecnone"">人材派遣</a>｜<a href=""" & HTTP_CURRENTURL & "s_contents/s_introduce.asp"" class=""topdecnone"">人材紹介</a>｜<a href=""" & HTTP_CURRENTURL & "s_contents/s_temptoperm.asp"" class=""topdecnone"">紹介予定派遣</a></li>"
	Response.Write "<li><a href=""" & HTTPS_CURRENTURL & "staff/access.asp"" class=""topdecnone"">お問合せ</a></li>"
	Response.Write "</ul>"

	Response.Write "<ul>"
	Response.Write "<li class=""ttl"">特集</li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "order/special/ad/0001/"" class=""topdecnone"">SE転職</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "order/special/tg/0004/"" class=""topdecnone"">臨床検査技師の求人</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "order/special/tg/0005/"" class=""topdecnone"">英語を活かして派遣で働く</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "order/special/sz/0001/"" class=""topdecnone"">静岡で転職!!</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "order/special/ng/0002/"" class=""topdecnone"">名古屋の派遣</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "order/special/or/0001/"" class=""topdecnone"">DTPオペレーター・デザイナー求人</a></li>"
	If Now <= "2011/09/15 12:00:00" Then
		'<キャンペーン>
		Response.Write "<li><a href=""" & HTTPS_CURRENTURL & "campaign/2011090101/"" target=""_blank"" class=""topdecnone"" style=""font-size:95%;"">岡山限定!営業職の転職支援強化ｷｬﾝﾍﾟｰﾝ</a></li>"
		'</キャンペーン>
	Else
		Response.Write "<li><a href=""" & HTTP_CURRENTURL & "order/special/oy/0001/"" class=""topdecnone"">岡山の求人</a></li>"
	End if
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "order/special/hr/0001/"" class=""topdecnone"">広島で転職!!</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "s_contents/license/1700101.asp"" class=""topdecnone"">宅地建物取引主任者 求人</a></li>"
	Response.Write "</ul>"

	Response.Write "<ul style=""margin-right:0px;"">"
	Response.Write "<li class=""ttl"">求人企業様</li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "tab/index5.asp"" class=""topdecnone"">採用ご担当者様</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "company/c_hajime.asp"" class=""topdecnone"">求人広告</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "company/c_staffdata.asp"" class=""topdecnone"">しごとナビ求職者と掲載企業データ</a></li>"
	Response.Write "<li><a href=""" & HTTP_CURRENTURL & "company/c_dispatch.asp"" class=""topdecnone"">人材派遣</a>｜<a href=""" & HTTP_CURRENTURL & "company/c_introduce.asp"" class=""topdecnone"">人材紹介</a>｜<a href=""" & HTTP_CURRENTURL & "company/c_temptoperm.asp"" class=""topdecnone"">紹介予定派遣</a></li>"
	Response.Write "<li><a href=""" & HTTPS_CURRENTURL & "company/access.asp"" class=""topdecnone"">お問合せ</a></li>"
'	Response.Write "<li><a href=""" & HTTPS_CURRENTURL & "company/access.asp"" class=""topdecnone"">広告代理店の方のお問合せ</a></li>"
'	Response.Write "<li><a href=""" & HTTPS_CURRENTURL & "company/fc_index.asp" class="topdecnone">しごとナビFC</a></li>"
	Response.Write "</ul>"

	Response.Write "<br clear=""all"">"

	Response.Write "<div style=""text-align:center;"">"
	Response.Write "<a href=""" & HTTP_LIS_CURRENTURL & """ target=""_blank""><img src=""/img/footer/footer_lis_logo_1.gif"" alt=""転職サイト｢しごとナビ｣運営-リス株式会社-"" border=""0""></a>"
	Response.Write "</div>"

	Response.Write "</div>"
	Response.Write "</div>"
	Response.Write "</div>" & vbCrLf

	'<Twitterバッジ>
'	Select Case getTabIndexType(Request.ServerVariables("URL"))
'		Case 0,1,2,3,4,6: Response.Write scrTwitterFollowBadge()
'		Case 5,7: Response.Write scrIntroTwitterFollowBadge()
'	End Select
	'</Twitterバッジ>

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
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_goukaku.asp"">合格率ＵＰマニュアル</a></li>"
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

	Response.Write "</div>"
End Function
%>
