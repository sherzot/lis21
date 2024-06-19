<%
'******************************************************************************
'概　要：ヘッダー
'引　数：HeadType	0【トップ】1【求職者】2【企業】3【共用】
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

	If HeadType = 0 Then 'トップ
		'sHeadcmt = "　転職活動に必要な書類(履歴書等)の作成・お仕事情報・プロによる貴方に適した転職サポートをご提供しています！"
		sHeadcmt = "<div style=""padding-left:8px; color:#ffffff;"">求人,募集情報はもちろん、履歴書・職務経歴書等の自動作成、プロによる貴方に適した転職サポートをご提供しています！</div>"
		sLinkurl = "/company/index.asp"
		sLinkalt = "求人広告しごとナビ"
		sLinktext = "採用担当（求人企業）様はこちら"
	ElseIf HeadType = 1 Then '求職者
		sHeadcmt = "<div style=""padding-left:8px;"">転職活動・求職活動の方々に最適な求人情報と履歴書ツールを提供しています</div>"
		sLinkurl = "/company/index.asp"
		sLinkalt = "求人広告しごとナビ"
		sLinktext = "採用担当（求人企業）様はこちら"
	ElseIf HeadType = 2 Then '企業
		sHeadcmt = "<div style=""padding-left:8px;"">企業の人材雇用を幅広くサポートしております。（求人広告、人材派遣、人材紹介）</div>"
		sLinkurl = "/"
		sLinkalt = "転職・求人サイトしごとナビ"
		sLinktext = "お仕事をお探しの方はこちら"
	ElseIf HeadType = 3 Then '共用
		sHeadcmt = "<div style=""padding-left:8px;"">求職活動の方々に最適な求人情報と履歴書ツールを提供しています</div>"
		sLinkurl = "/company/index.asp"
		sLinkalt = "求人広告しごとナビ"
		sLinktext = "採用担当（求人企業）様はこちら"
	End If

	Response.Write "<div id=""wrap"" align=""center"">"
	Response.Write "<div id=""wrapw"">"
	Response.Write "<div id=""head"" align=""left"">"
	Response.Write "<table>"
	Response.Write "<tr>"

	'<ヘッダー左：しごとナビロゴ>
	Response.Write "<td align=""left"" style=""height:42px; width:141px;"">"
	'クリスマス
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
		Response.Write "<a class=""decnone"" href=""/"" title=""転職・求人サイト「しごとナビ」""><img src=""/img/top/shigotonavi_logo.gif"" alt=""しごとナビ"" border=""0"" align=""left"" style=""margin-left:4px;""></a>"
	End If

	Response.Write "</td>"
	'</ヘッダー左：しごとナビロゴ>

	'<ヘッダー右>
	Response.Write "<td align=""right"" style=""font-size:11px;"">"
	Response.Write "　<a href=""/staff/access.asp""><img src=""/img/top/head_icon.gif"" alt=""お問合せ"" border=""0"" style=""vertical-align:millde;"">お問合せ</a>"
	Response.Write "　<a href=""" & HTTP_CURRENTURL & "shigotonavi/sitemap.asp""><img src=""/img/top/head_icon.gif"" alt=""サイトマップ"" border=""0"" style=""vertical-align:middle;"">サイトマップ</a>"

	'<Googleのサイト内検索>
	If Request.ServerVariables("HTTPS") <> "on" Then
		Response.Write "<form action=""/search.asp"" id=""cse-search-box"" style=""margin-left:5px;padding:0px;display:inline"">"
		Response.Write "<div style=""display:inline;"">"
		Response.Write "<img src=""/img/top/head_icon.gif"" alt="""" border=""0"" style=""vertical-align:millde;"">"
		Response.Write "<label>"
		Response.Write "<span>サイト内検索&nbsp;</span>"
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
	'</Googleのサイト内検索>

	Response.Write "　<a href=""" & sLinkurl & """ title=""" & sLinkalt & """ style=""font-size:14px;""><img src=""/img/top/head_icon.gif"" alt=""" & sLinkalt & """ border=""0"" style=""vertical-align:middle;"">" & sLinktext & "</a>"
	'<!-- #INCLUDE FILE="../ad_banner_control/ad_banner.asp" -->
	Response.Write "</td>"
	'<ヘッダー右>

	Response.Write "</tr>"

	'<ヘッダー下部：背景緑のやつ>
	Response.Write "<tr style=""background-image:url(/img/top/headtext_background.gif);"">"
	Response.Write "<td colspan=""2"" align=""left"" style=""margin:0px;padding:0px;color:#ffffff; height:20px;border-top:solid 1px #ffffff; border-bottom:solid 1px #ffffff;"">"
	Response.Write sHeadcmt
	Response.Write "</td>"
	Response.Write "</tr>"
	'</ヘッダー下部：背景緑のやつ>

	Response.Write "</table>"
	Response.Write "</div>"
	Response.Write "<div align=""left"" style=""width:100%;background-color:#ffffff;"">"
	Response.Write "<div align=""left"" style=""width:790px;float:left;"">" 'ページ全体の幅（footer最下部で閉め
	Response.Write "<div class=""moji912"" style=""padding-left:3px;width:615px;float:right"">" & vbCrLf 'メインコンテンツ幅（sidemenu最上部で閉め）
End Function

'******************************************************************************
'概　要：サイドメニュー
'引　数：SidemenuType	0【トップ】1【求職者】2【企業】3【共用】
'作成者：Lis Niina
'作成日：2008/02/07
'備　考：
'使用元：
'履　歴：
' 08/05/20 Lis林 しごとナビFC追加
'******************************************************************************
Function NaviSidemenu(SidemenuType)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	If SidemenuType = 0 Then 'トップページ
		Response.Write "</div>"'メインコンテンツの幅指定divの閉め（開始はheader最下部）
		Response.Write "<div style=""width:170px; float:left; margin:0px;padding:0px;"">"

		'▽トップページもログイン情報によってサイドメニューを切り替える
		If session("usertype") = "staff" Then '求職者ログインしている場合
			Response.Write "<ul>"
			Response.Write "<li class=""sidemenu_staff_big"">My Menu （<a title=""ログアウト"" href=""" & HTTP_CURRENTURL & "logout.asp"" style=""font-size:11px;"">ログアウトする</a>）</li>"
			Response.Write "<li class=""sidemenu_mypage""><a title=""My Page"" href=""" & HTTPS_CURRENTURL & "login_menu.asp"">My Page</a></li>"
			Response.Write "<li class=""sidemenu_job""><a title=""ジョブ・コンシェルジュ"" href=""" & HTTP_CURRENTURL & "staff/jobcon/"">ジョブ・コンシェルジュ</a></li>"
			Response.Write "<li class=""sidemenu_job""><a title=""お仕事検索"" href=""" & HTTP_CURRENTURL & "order/order_search_detail.asp"">お仕事検索</a></li>"
			Response.Write "<li class=""sidemenu_mail""><a title=""メール管理"" href=""" & HTTPS_CURRENTURL & "staff/mailhistory_person.asp"">メール管理"

			sSQL = "SELECT COUNT(*) AS Cnt FROM MailHistory WITH(NOLOCK) WHERE ReceiverCode ='" & Session("userid") & "' AND OpenDay IS NULL AND ReceiverDelFlag = '0'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				If oRS.Collect("Cnt") = 0 Then
					Response.Write "(<img src=""/img/staff/mail/mailhei.gif"" border=""0"" alt="""" style=""margin:0px 1px;"">未読" & oRS.Collect("Cnt") & "件)"
				Else
					Response.Write "(<span style=""color:#ff0000; font-weight:bold;""><img src=""/img/staff/mail/mailhei.gif"" border=""0"" alt="""" style=""margin:0px 1px;"">未読" & oRS.Collect("Cnt") & "件</span>)"
				End If
			End If

			Response.Write "</a></li>"
			Response.Write "<li class=""sidemenu_detail""><a title=""登録内容修正"" href=""" & HTTPS_CURRENTURL & "staff/person_detail.asp"">登録内容修正</a></li>"
			Response.Write "<li class=""sidemenu_print""><a title=""履歴書・職務経歴書　印刷"" href=""" & HTTP_CURRENTURL & "staff/resume_print.asp"">履歴書・職務経歴書　出力</a></li>"
			Response.Write "<li class=""sidemenu_wacth""><a title=""ウォッチリスト"" href=""" & HTTP_CURRENTURL & "staff/watchlist.asp"">ウォッチリスト</a></li>"
			Response.Write "<li class=""sidemenu_picture""><a title=""履歴書写真登録"" href=""" & HTTP_CURRENTURL & "staff/resume_picture.asp"">履歴書写真登録</a></li>"
			Response.Write "<li class=""sidemenu_pass""><a title=""パスワードの変更"" href=""" & HTTPS_CURRENTURL & "staff/changepassword.asp"">パスワード変更</a></li>"
			Response.Write "<li class=""sidemenu_staff_bottom""></li>"
			Response.Write "</ul>"
		ElseIf ( Session("usertype") = "company" Or Session("usertype") = "dispatch") And G_USEFLAG <> "0" Then '企業ログインしている場合
			Response.Write "<ul>"
			Response.Write "<li class=""sidemenu_company_big"">My Menu （<a title=""ログアウト"" href=""" & HTTP_CURRENTURL & "logout.asp"" style=""font-size:11px;"">ログアウトする</a>）</li>"
			Response.Write "<li class=""sidemenu_company""><a title=""My Page"" href=""" & HTTPS_CURRENTURL & "login_menu.asp"">My Page</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""求職者の検索"" href=""" & HTTP_CURRENTURL & "company/myorderlist.asp"">求職者の検索</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""ウォッチリスト"" href=""" & HTTP_CURRENTURL & "company/watchlist.asp"">ウォッチリスト</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""メール履歴"" href=""" & HTTPS_CURRENTURL & "company/mailhistory_company.asp"">メール履歴</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""求人票の修正"" href=""" & HTTP_CURRENTURL & "company/myorderlist.asp"">求人票の修正</a></li>"

			If Session("usertype") = "company" Then
				Response.Write "<li class=""sidemenu_company""><a href=""" & HTTPS_CURRENTURL & "company/company_reg1.asp"">自社情報を更新</a></li>"
				If G_IMAGELIMIT > 0 Then
					Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "company/img_upload.asp"">企業写真画像掲載</a></li>"
				End If

				If G_IMAGELIMIT > 1 Then
					Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "company/company_img_list.asp"">求人票用画像ストック</a></li>"
				End If
			ElseIf Session("usertype") = "dispatch" Then
				Response.Write "<li class=""sidemenu_company""><a href=""" & HTTPS_CURRENTURL & "dispatch/company_reg1.asp"">自社情報を更新</a></li>"
				If G_IMAGELIMIT > 0 Then
					Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "company/img_upload.asp"">企業写真画像掲載</a></li>"
				End If

				If G_IMAGELIMIT > 1 Then
					Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "company/company_img_list.asp"">求人票用画像ストック</a></li>"
				End If
			End If

			Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "mailtemplate/manager.asp"">メールテンプレート管理</a></li>"
			If G_PLANTYPE <> "mail" then
				Response.Write "<li class=""sidemenu_company""><a href=""" & HTTPS_CURRENTURL & "company/costperformance/"">採用改善ｻﾎﾟｰﾄｼｽﾃﾑ<img src=""/img/new.gif"" border=""0""></a></li>"
			End If

			'Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "license/license_manager.asp"">ライセンス管理</a></li>"
			Response.Write "<li class=""sidemenu_company_end""><a href=""" & HTTPS_CURRENTURL & "company/changepassword.asp"">パスワード変更</a></li>"
			Response.Write "<li class=""sidemenu_company_bottom""></li>"
			Response.Write "</ul>"
		ElseIf (Session("usertype") = "company" Or Session("usertype") = "dispatch") And G_USEFLAG = "0"  Then '企業ログインしているがライセンスが切れている場合
			Response.Write "<ul>"
			Response.Write "<li class=""sidemenu_company_big"">My Menu</li>"
			Response.Write "<li class=""sidemenu_company""><a title=""My Page"" href=""" & HTTPS_CURRENTURL & "login_menu.asp"">My Page （<a title=""ログアウト"" href=""" & HTTP_CURRENTURL & "logout.asp"" style=""font-size:11px;"">ログアウトする</a>）</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""求職者の検索"" href=""" & HTTP_CURRENTURL & "company/myorderlist.asp"">求職者の検索</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""ウォッチリスト"" href=""" & HTTP_CURRENTURL & "company/watchlist.asp"">ウォッチリスト</a></li>"

			If G_MAILREADFLAG = "1" Then
				Response.Write "<li class=""sidemenu_company""><a title=""メール履歴"" href=""" & HTTPS_CURRENTURL & "company/mailhistory_company.asp"">メール履歴</a></li>"
			End If

			Response.Write "<li class=""sidemenu_company""><a title=""自社求人票一覧"" href=""" & HTTP_CURRENTURL & "company/myorderlist.asp"">自社求人票一覧</a></li>"
			'Response.Write "<li class=""sidemenu_company""><a title=""ライセンス管理"" href=""" & HTTP_CURRENTURL & "license/license_manager.asp"">ライセンス管理</a></li>"
			Response.Write "<li class=""sidemenu_company_end""><a title=""パスワード変更"" href=""" & HTTPS_CURRENTURL & "company/changepassword.asp"">パスワード変更</a></li>"
			Response.Write "<li class=""sidemenu_company_bottom""></li>"
			Response.Write "</ul>"

			'Response.Write "<ul>"
			'Response.Write "<li class=""sidemenu_big"">ログイン<span style=""font-size:10px;"">（既にログイン済みです）</span></li>"
			'Response.Write "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "login_menu.asp"" title=""転職ログイン"">My Pageへ</a></li>"
			'Response.Write "<li class=""sidemenu_bottom""></li>"
			'Response.Write "</ul>"
			'△トップページもログイン情報によってサイドメニューを切り替える
		Else
			Response.Write "<div align=""center"">"
			Response.Write "<a href=""" & HTTPS_CURRENTURL & "staff/person_reg1.asp"" title=""転職,新規会員登録""><img src=""/img/common/reg1_button.jpg"" border=""0"" alt=""新規会員登録"" style=""margin-top:3px;""></a>"
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
			Response.Write "<li class=""sidemenu_big"">ログイン</li>"
			Response.Write "<li style=""border-right:solid 1px #cccccc; border-left:solid 1px #cccccc;"">"
			Response.Write "<div style=""font-size:11px; padding-top:0px; padding-right:3px;"">"
			Response.Write "<div align=""right"">"

			If G_SSLFLAG = False Then
				Response.Write "<a href=""" & HTTPS_CURRENTURL & """ style=""color:#0045f9;""><img src=""/img/common/security_key.gif"" border=""0"" height=""12"" alt="""">ＳＳＬをＯＮにする ※推奨</a><br>"
			Else
				Response.Write "<a href=""" & HTTP_CURRENTURL & """ style=""color:#0045f9;"">ＳＳＬをＯＦＦにする</a><br>"
			End If
			Response.Write "</div>"

			If si <> "" Then
				Response.Write "<p class=""m0"" style=""float:right;"">　<input type=""text"" name=""CONF_UserID"" value=""" & si & """ style=""width:80px;""></p>"
			Else
				Response.Write "<p class=""m0"" style=""float:right;"">　<input type=""text"" name=""CONF_UserID"" value=""" & Request.Cookies("id_memory") & """ style=""width:80px;""></p>"
			End If

			Response.Write "<p class=""m0"" style=""font-size:10px;color:#666666;float:right;""><b>I　D</b></p>"
			Response.Write "<br clear=""all"">"
			Response.Write "<p class=""m0"" style=""float:right;"">　<input type=""password"" name=""CONF_Password"" value="""" style=""width:80px;""></p>"
			Response.Write "<p class=""m0"" style=""font-size:10px;color:#666666;float:right;""><b>パスワード</b></p>"
			Response.Write "<br clear=""all"">"
			Response.Write "<div align=""right"">"
			Response.Write "<label><input type=""checkbox"" name=""frmautologinflag"" value=""1"">自動ﾛｸﾞｲﾝ</label>[<span style=""color:#0045f9; cursor:pointer;"" onclick=""window.open('/infomation/autologin.asp','autologin','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=400,height=220');""><u>？</u></span>]"
			Response.Write "<input type=""submit"" value=""ログイン"" onclick=""DataCheckIdreg(); return false""><br>"
			Response.Write "<a href=""staff/qa.asp#003"" style=""font-size:10px;"" title=""転職,ログインできない方"">ﾛｸﾞｲﾝできない</a>　"
			Response.Write "<a href=""" & HTTPS_CURRENTURL & "staff/passwordreminder.asp"" style=""font-size:10px;"" title=""転職,パスワードを忘れた方"">ID・ﾊﾟｽﾜｰﾄﾞを忘れた</a><br>"
			Response.Write "</div>"
			Response.Write "</div>"
			Response.Write "</li>"
			Response.Write "<li class=""sidemenu_bottom""></li>"
			Response.Write "</ul>"
%><!-- #INCLUDE FILE="../error/errHandle.asp" --><%
			Response.Write "</form>"
		End If

		'トップページ左
		'Response.Write "<div style=""width:170px;height:50px;margin-bottom:5px;"">"
		'Response.Write "<a href=""" & HTTP_CURRENTURL & "order/order_detail.asp?OrderCode=J0051817"" title=""SOHO広告代理店"">"
		'Response.Write "<img src=""/img/top/soho_banner.gif"" alt=""SOHO広告代理店"" border=""0""><br>"
		'Response.Write "</a>"
		'Response.Write "</div>"

		Response.Write "<div style=""width:170px;height:50px;margin-bottom:5px;"">"
		Response.Write "<a href=""" & HTTP_CURRENTURL & "staff/jobcon/introduction.asp"" title=""ジョブ・コンシェルジュ"">"
		Response.Write "<img src=""/img/staff/jobcon/top_mini_banner.gif"" alt=""転職支援ジョブ・コンシュルジュ"" border=""0""><br>"
		Response.Write "</a>"
		Response.Write "</div>"
		Response.Write "<div style=""width:170px;height:135px;background-image:url(/img/sidemenu/navicafe_banner_all.jpg);margin-bottom:5px;"">"
		Response.Write "<a href=""/cafe/cafe_list.asp"" title=""ナビカフェ""><img src=""/img/sidemenu/navicafe_banner_top.jpg"" alt=""ナビカフェ"" border=""0"" style=""margin:0px;padding:0px;""></a>"
		Response.Write "<div style=""margin-top:14px;padding:0px 6px 0px 8px;font-size:10px;line-height:15px;"">"

		'** TOP 08/11/05 Lis林 ADD
		'現在掲載中＆TOP3のトピ
		sSQL = "EXEC up_GetData_NC_Topic '','','','1','3';"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		Do While GetRSState(oRS) = True
			Response.Write "<a href='/cafe/cafe_detail.asp?t=" & oRS.Collect("TopicID")
			Response.Write "' title='" & oRS.Collect("Title") & "'>・"
			If Len(oRS.Collect("Title")) > 14 Then
				Response.Write Left(oRS.Collect("Title"),14) & "..."
			Else
				Response.Write oRS.Collect("Title")
			End If
			Response.Write "</a><br>"
			oRS.MoveNext
		Loop
		Call RSClose(oRS)
		'** BTM 08/11/05 Lis林 ADD

		Response.Write "</div>"
		Response.Write "</div>"

		Response.Write "<div style=""width:170px;height:135px;background-image:url(/img/sidemenu/warmreception_banner_all.jpg);margin-bottom:5px;"">"
		Response.Write "<a href=""/s_contents/warmreception/"" title=""しごとナビ会員優待""><img src=""/img/sidemenu/warmreception_banner_top.jpg"" alt=""しごとナビ会員優待"" border=""0""></a>"
		Response.Write "<div style=""margin-top:14px;padding:0px 6px 0px 8px;font-size:10px;line-height:15px;"">"
		Response.Write "<a href=""" & HTTP_CURRENTURL & "s_contents/warmreception/detail.asp?category=license&id=0102"">・ビジネス・キャリア検定試験（２級）</a><br>"
		Response.Write "<a href=""" & HTTP_CURRENTURL & "s_contents/warmreception/detail.asp?category=license&id=0103"">・ビジネス・キャリア検定試験（３級）</a><br>"
		Response.Write "<a href=""" & HTTP_CURRENTURL & "s_contents/warmreception/detail.asp?category=skillup&id=0101"">・語学スキル 英会話（マンツーマン）</a><br>"
		Response.Write "</div>"
		Response.Write "</div>"

		Response.Write "<ul>"
		Response.Write "<li class=""sidemenu_big"">便利ツール</li>"
		Response.Write "<li style=""border-left:solid 1px #cccccc; border-right:solid 1px #cccccc; border-bottom:solid 1px #dddddd; line-height:17px;""><a href=""/staff/s_resume.asp"" title=""履歴書の自動作成"" style=""display:block; background-image:url(/img/sidemenu/resume_banner.jpg); width:154px; height:73px; font-size:10px; padding:54px 0px 0px 14px; color:#444444; text-decoration:none;"">必要な項目を入力するだけで完成！<br>" & Left(sAll,2) & "万人が使う安心のサービス！<br>自分に合った履歴書が作れる！</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/staff/s_resume_kakikata.asp"" title=""履歴書の書き方"">履歴書の書き方</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/staff/s_careersheet.asp"" title=""職務経歴書の自動作成"">職務経歴書の自動作成</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/staff/s_careersheet_kakikata_1.asp"" title=""職務経歴書の書き方"">職務経歴書の書き方</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/s_contents/motive_index.asp"" title=""志望動機メーカー"">志望動機メーカー</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/s_contents/s_jikopr.asp"" title=""自己PRメーカー"">自己PRメーカー</a></li>"
		Response.Write "<li class=""sidemenu_end""><a href=""/s_contents/s_taishokunegai.asp"" title=""退職願の書き方"">退職願の書き方</a></li>"
		Response.Write "<li class=""sidemenu_bottom""></li>"
		Response.Write "</ul>"

		Response.Write "<ul>"
		Response.Write "<li class=""sidemenu_big"">サポート</li>"
		Response.Write "<li class=""sidemenu""><a href=""/s_contents/navistep_index.asp"" title=""初めての転職活動"">初めての転職活動</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/staff/s_aboutnavi.asp"" title=""ご利用ガイド"">ご利用ガイド</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/staff/qa.asp"" title=""Ｑ＆Ａ"">Ｑ＆Ａ</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/staff/s_searchexplanation.asp"" title=""お仕事検索方法"">お仕事検索方法</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/staff/s_kiyaku.asp"" title=""利用規約"">利用規約</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/shigotonavi/sitemap.asp"" title=""サイトマップ"">サイトマップ</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/link.asp"" title=""リンクポリシー"">リンクポリシー</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/link_collection.asp"" title=""お役立ち厳選リンク集"">お役立ち厳選リンク集</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/s_contents/s_books.asp"" title=""転職に役立つ本"">転職に役立つ本</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/company/index.asp"" title=""企業向け採用コンテンツ"">企業向け求人広告について</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""/lis/lis.asp"" title=""運営会社"">運営会社</a></li>"
		Response.Write "<li class=""sidemenu_end""><a href=""" & HTTPS_CURRENTURL & "staff/access.asp"" title=""お問い合わせ"">お問い合わせ</a></li>"
		Response.Write "<li class=""sidemenu_bottom""></li>"
		Response.Write "</ul>"

		Response.Write "<div align=""center"" style=""width:100%;"">"
		Response.Write "<div class=""sidemenu_big"" style=""text-align:left;"">求職者情報</div>"
		Response.Write "<div style=""border-left:solid 1px #cccccc; border-right:solid 1px #cccccc; background-image:url(/img/sidemenu/jinzaidata_background.gif);"" align=""center"">"
		Response.Write "<table style=""width:155px; font-size:10px; text-align:left;"">"

		Dim rank(2)
		Dim rankcount(2)
		Dim idx
		idx = 0

		sSQL = "SELECT top 3 Subitem,Number FROM Person_Statistics where item = '都道府県別' order by convert(int,Number) desc"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		Do While GetRSState(oRS) = True
			rank(idx) = Replace(Replace(Replace(oRS.Collect("SubItem"),"都",""),"府",""),"県","")
			rankcount(idx) = oRS.Collect("Number")
			idx = idx + 1
			oRS.MoveNext
		Loop
		Call RSClose(oRS)

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
		sSQL = "SELECT top 3 item,subitem, Number FROM Person_Statistics where item = '10歳代' or item = '20歳代' or item = '30歳代' or item = '40歳代' or item = '50歳代' or item = '60歳以上' order by convert(int,Number) desc"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		Do While GetRSState(oRS) = True
			rank(idx) = Replace(oRS.Fields("Item").Value,"歳","") & oRS.Fields("SubItem").Value
			rankcount(idx) = oRS.Fields("Number").Value
			idx = idx + 1
			oRS.MoveNext
		Loop
		Call RSClose(oRS)

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
		Response.Write "<td>3位:" & rank(2) & "</td>"
		Response.Write "<td align=""right"">" & rankcount(2) & "名</td>"
		Response.Write "</tr>"
		Response.Write "<tr>"
		Response.Write "<td colspan=""3"" align=""right""><a href=""/company/c_staffdata.asp""><img src=""/img/sidemenu/kuwashiku_min.jpg"" alt=""詳しくはこちら"" border=""0""></a>"
		Response.Write "</tr>"
		Response.Write "</table>"
		Response.Write "</div>"
		Response.Write "<div class=""sidemenu_bottom"" style=""clear:both;""></div>"
		Response.Write "<br>"
		Response.Write "</div>"

		Response.Write "<div align=""center"" style=""width:100%;"">"
		Response.Write "<a href=""/lis/blog_kimura.asp"">"
		Response.Write "<img src=""/img/top/top_blogBanner.gif"" border=""0"" alt=""木村亮郎のヒトビジネスつれづれ"">"
		Response.Write "</a>"
		Response.Write "</div>"

		'Response.Write "<div style=""text-align:center; font-size:11px;width:100%;padding:0px 15px;"">"
		'Response.Write "<img src=""/img/spacer.gif"" width=""3"" height=""10"" alt=""転職""><br>"
		'Response.Write "<div style=""float:left;"">"
		'Response.Write "<a href=""http://privacymark.jp/"" target=""_blank""><img src=""/img/privacy/p_75.gif"" alt=""プライバシーマーク"" border=""0"" width=""45""></a><br><a href=""/privacy/privacy.asp"">個人情報保護</a></div><div>"
		'Response.Write "<a href=""https://secure.secom.ne.jp/webp/db/1116062419.html"" target=""_blank""><img src=""img/secom/B0474507/B0474507_s.gif"" border=""0"" alt="""" height=""43""><br>SSL暗号化<br></a>"
		'Response.Write "</div>"
		'Response.Write "</div>"

		Response.Write "<div style=""text-align:center""></div>"
		Response.Write "</div>"
	ElseIf SidemenuType = 1 Then '求職者
		Response.Write "</div>" 'メインコンテンツの幅指定divの閉め（開始はheader最下部）
		Response.Write "<div id=""idNavigation"" style=""width:170px; float:left;"">"

		Response.Write "<!-- MENU START -->"
		'■■ログイン時の求職者左側
		If Session("usertype") = "staff" Then
			Response.Write "<div style=""clear:both; margin-bottom:5px;""></div>"
			Response.Write "<ul>"
			Response.Write "<li class=""sidemenu_staff_big"">My Menu （<a title=""ログアウト"" href=""" & HTTP_CURRENTURL & "logout.asp"" style=""font-size:11px;"">ログアウトする</a>）</li>"
			Response.Write "<li class=""sidemenu_mypage""><a title=""My Page"" href=""" & HTTPS_CURRENTURL & "login_menu.asp"">My Page</a></li>"
			Response.Write "<li class=""sidemenu_job""><a title=""ジョブ・コンシェルジュ"" href=""" & HTTP_CURRENTURL & "staff/jobcon/"">ジョブ・コンシェルジュ</a></li>"
			Response.Write "<li class=""sidemenu_job""><a title=""お仕事検索"" href=""" & HTTP_CURRENTURL & "order/order_search_detail.asp"">お仕事検索</a></li>"
			Response.Write "<li class=""sidemenu_mail""><a title=""メール管理"" href=""" & HTTPS_CURRENTURL & "staff/mailhistory_person.asp"">メール管理"

			sSQL = "SELECT COUNT(*) AS Cnt FROM MailHistory WITH(NOLOCK) WHERE ReceiverCode ='" & Session("userid") & "' AND OpenDay IS NULL AND ReceiverDelFlag = '0'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

			If GetRSState(oRS) = True Then
				If oRS.Collect("Cnt") = 0 Then
					Response.Write "(<img src=""/img/staff/mail/mailhei.gif"" border=""0"" alt="""" style=""margin:0px 1px;"">未読" & oRS.Collect("Cnt") & "件)"
				Else
					Response.Write "(<span style=""color:#ff0000; font-weight:bold;""><img src=""/img/staff/mail/mailhei.gif"" border=""0"" alt="""" style=""margin:0px 1px;"">未読" & oRS.Collect("Cnt") & "件</span>)"
				End If
			End If

			Response.Write "</a></li>"
			Response.Write "<li class=""sidemenu_detail""><a title=""登録内容修正"" href=""" & HTTPS_CURRENTURL & "staff/person_detail.asp"">登録内容修正</a></li>"
			Response.Write "<li class=""sidemenu_print""><a title=""履歴書・職務経歴書　印刷"" href=""" & HTTP_CURRENTURL & "staff/resume_print.asp"">履歴書・職務経歴書　出力</a></li>"
			Response.Write "<li class=""sidemenu_wacth""><a title=""ウォッチリスト"" href=""" & HTTP_CURRENTURL & "staff/watchlist.asp"">ウォッチリスト</a></li>"
			Response.Write "<li class=""sidemenu_footprint""><a title=""気になリスト"" href=""" & HTTP_CURRENTURL & "staff/footprint.asp"">気になリスト</a></li>"
			Response.Write "<li class=""sidemenu_picture""><a title=""履歴書写真登録"" href=""" & HTTP_CURRENTURL & "staff/resume_picture.asp"">履歴書写真登録</a></li>"
			Response.Write "<li class=""sidemenu_pass""><a title=""パスワードの変更"" href=""" & HTTPS_CURRENTURL & "staff/changepassword.asp"">パスワード変更</a></li>"
			Response.Write "<li class=""sidemenu_staff_bottom""></li>"
		Else
			'■■ログインしていない時の求職者左側
			Response.Write "<a href=""" & HTTPS_CURRENTURL & "staff/person_reg1.asp""><img src=""/img/common/reg1_button.jpg"" alt=""しごとナビ会員登録"" border=""0"" style=""margin:3px 0px 2px 0px;""></a><br>"
			Response.Write "<div align=""right"" style=""font-size:11px; margin-bottom:5px;"">"
			Response.Write "<a href=""" & HTTPS_CURRENTURL & "login_menu.asp"">会員登録がお済みの方はこちら</a>"
			Response.Write "</div>"
			Response.Write "<a title=""お仕事検索"" href=""" & HTTP_CURRENTURL & "order/order_search_detail.asp""><img src=""/img/sidemenu/jobsearch_button.jpg"" alt=""お仕事検索"" border=""0"" style=""margin-top:3px;""></a>"
			Response.Write "<div style=""width:170px;height:50px;margin:5px 0px;"">"
			Response.Write "<a href=""" & HTTP_CURRENTURL & "staff/jobcon/introduction.asp"" title=""ジョブ・コンシェルジュ"">"
			Response.Write "<img src=""/img/staff/jobcon/top_mini_banner.gif"" alt=""転職支援ジョブ・コンシュルジュ"" border=""0""><br>"
			Response.Write "</a>"
			Response.Write "</div>"

			Response.Write "<ul>"
		End If

		'■■求職者左側共通
		Response.Write "<li class=""sidemenu_big"">コミュニティ</li>"
		Response.Write "<li class=""sidemenu""><a title=""ナビカフェ"" href=""" & HTTP_CURRENTURL & "cafe/cafe_list.asp"">ナビカフェ</a></li>"
		Response.Write "<li class=""sidemenu_end""><a title=""しごとナビアンケート"" href=""" & HTTP_CURRENTURL & "s_contents/enquete.asp"">しごとナビアンケート</a></li>"
		Response.Write "<li class=""sidemenu_bottom""></li>"

		Response.Write "<li class=""sidemenu_big"">書類作成支援</li>"
		Response.Write "<li class=""sidemenu""><a title=""履歴書の自動作成"" href=""" & HTTP_CURRENTURL & "staff/s_resume.asp"">履歴書の自動作成</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""履歴書の書き方"" href=""" & HTTP_CURRENTURL & "staff/s_resume_kakikata.asp"">履歴書の書き方</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""履歴書Ｑ＆Ａ"" href=""" & HTTP_CURRENTURL & "staff/s_resume_qa.asp"">履歴書Ｑ＆Ａ</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""職務経歴書の自動作成"" href=""" & HTTP_CURRENTURL & "staff/s_careersheet.asp"">職務経歴書の自動作成</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""職務経歴書の書き方"" href=""" & HTTP_CURRENTURL & "staff/s_careersheet_kakikata_1.asp"">職務経歴書の書き方</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""志望動機メーカー"" href=""" & HTTP_CURRENTURL & "s_contents/motive_index.asp"">志望動機メーカー</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""自己PRメーカー"" href=""" & HTTP_CURRENTURL & "s_contents/s_jikopr.asp"">自己PRメーカー</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""退職願の書き方"" href=""" & HTTP_CURRENTURL & "s_contents/s_taishokunegai.asp"">退職願の書き方</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""学歴計算・西暦和暦早見表"" href=""" & HTTP_CURRENTURL & "s_contents/s_year_calculation.asp"">学歴計算・西暦和暦早見表</a></li>"
		Response.Write "<li class=""sidemenu_end""><a title=""Conpri - コンプリ"" href=""" & HTTP_CURRENTURL & "conpri/"">コンビニ印刷</a></li>"
		Response.Write "<li class=""sidemenu_bottom""></li>"

		Response.Write "<li class=""sidemenu_big"">転職支援ツール</li>"
		Response.Write "<li class=""sidemenu""><a title=""初めての転職活動"" href=""" & HTTP_CURRENTURL & "s_contents/navistep_index.asp"">初めての転職活動</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""しごとナビ転職コラム"" href=""" & HTTP_CURRENTURL & "column/column_index.asp"">しごとナビ転職コラム</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""転職の心構え"" href=""" & HTTP_CURRENTURL & "s_contents/s_ready.asp"">転職の心構え</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""転職に必要な手続き"" href=""" & HTTP_CURRENTURL & "s_contents/s_proce.asp"">転職に必要な手続き</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""合格率UP転職マニュアル"" href=""" & HTTP_CURRENTURL & "s_contents/s_goukaku.asp"">合格率UP転職マニュアル</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""ニートからの脱出"" href=""" & HTTP_CURRENTURL & "s_contents/s_neet.asp"">ニートからの脱出</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""紹介予定派遣とは"" href=""" & HTTP_CURRENTURL & "s_contents/s_temptoperm.asp"">紹介予定派遣とは</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""あなたの給与明細"" href=""" & HTTP_CURRENTURL & "s_contents/s_kyuuyomeisai.asp"">あなたの給与明細</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""適職診断｢じぶんナビ｣"" href=""" & HTTP_CURRENTURL & "s_contents/s_mynavi.asp"">適職診断｢じぶんナビ｣</a></li>"
		Response.Write "<li class=""sidemenu""><a title=""スカウトメールをたくさん受けるには！？"" href=""" & HTTP_CURRENTURL & "s_contents/labo/scoutlabo.asp"
		If G_USERTYPE = "staff" Then Response.Write "?staffcode=" & G_USERID & "&amp;linkno=3"
		Response.Write """>スカウトラボ</a></li>"

		If G_USERTYPE = "staff" Then
			Response.Write "<li class=""sidemenu_end""><a title=""しごとナビ会員優待"" href=""" & HTTP_CURRENTURL & "s_contents/warmreception/"">ナビ会員優待</a></li>"
		End If

		Response.Write "<li class=""sidemenu_bottom""></li>"


		Response.Write "</ul>"
		Response.Write "<br clear=""all"">"
		Response.Write "<div align=""center"" style=""clear:both; margin-top:20px;"">"
		Response.Write "<a href=""http://privacymark.jp/"" target=""_blank""><img src=""/img/privacy/p_75.gif"" alt=""プライバシーマーク"" border=""0""></a><br>"
		Response.Write "<a href=""/privacy/privacy.asp"">個人情報保護</a>"
		Response.Write "</div>"
		Response.Write "<div style=""text-align:center""></div>"
		Response.Write "<!-- MENU END -->"
		Response.Write "</div>"
	ElseIf SidemenuType = 2 Then '企業
		Response.Write "<script type=""text/javascript"">"
		Response.Write "function LoginCheckIdreg(){"
		Response.Write "var ofrm = document.forms.frmlogin;"
		Response.Write "if(!navigator.cookieEnabled) {"
		Response.Write "alert('cookie（クッキー）の利用ができない設定になっています。\nブラウザやセキュリティーソフトのcookie設定をご確認下さい。');"
		Response.Write "return false;"
		Response.Write "}"
		Response.Write "if(!ChkInput(ofrm.CONF_UserID, 'string', '1', '認証IDを入力してください。')) return false;"
		Response.Write "if(!ChkInput(ofrm.CONF_Password,'string', '1', 'パスワードを入力してください。')) return false;"
		Response.Write "if(!ChkLength(ofrm.CONF_Password, 3, 20, 'パスワードは３文字以上、２０文字以下で入力してください。'))return false;"
		Response.Write "ofrm.submit();"
		Response.Write "}"
		Response.Write "</script>"
		Response.Write "</div>"'メインコンテンツの幅指定divの閉め（開始はheader最下部）

		Response.Write "<div id=""idNavigation"" style=""width: 170px; float: left;"">"
		If G_USEFLAG = "0" Then
			'ライセンス切れの企業
			Response.Write "<ul>"
			Response.Write "<li class=""sidemenu_company_big"">My Menu （<a title=""ログアウト"" href=""" & HTTP_CURRENTURL & "logout.asp"" style=""font-size:11px;"">ログアウトする</a>）</li>"
			Response.Write "<li class=""sidemenu_company""><a title=""My Page"" href=""" & HTTPS_CURRENTURL & "login_menu.asp"">My Page</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""求職者の検索"" href=""" & HTTP_CURRENTURL & "company/myorderlist.asp"">求職者の検索</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""ウォッチリスト"" href=""" & HTTP_CURRENTURL & "company/watchlist.asp"">ウォッチリスト</a></li>"
			If G_MAILREADFLAG = "1" Then
				Response.Write "<li class=""sidemenu_company""><a title=""メール履歴"" href=""" & HTTPS_CURRENTURL & "company/mailhistory_company.asp"">メール履歴</a></li>"
			End If
			Response.Write "<li class=""sidemenu_company""><a title=""自社求人票一覧"" href=""" & HTTP_CURRENTURL & "company/myorderlist.asp"">自社求人票一覧</a></li>"
			'Response.Write "<li class=""sidemenu_company""><a title=""ライセンス管理"" href=""" & HTTP_CURRENTURL & "license/license_manager.asp"">ライセンス管理</a></li>"
			Response.Write "<li class=""sidemenu_company_end""><a title=""パスワード変更"" href=""" & HTTPS_CURRENTURL & "company/changepassword.asp"">パスワード変更</a></li>"
			Response.Write "<li class=""sidemenu_company_bottom""></li>"
		ElseIf Session("usertype") = "company" Or Session("usertype") = "dispatch" Then
			Response.Write "<ul>"
			Response.Write "<li class=""sidemenu_company_big"">My Menu （<a title=""ログアウト"" href=""" & HTTP_CURRENTURL & "logout.asp"" style=""font-size:11px;"">ログアウトする</a>）</li>"
			Response.Write "<li class=""sidemenu_company""><a title=""My Page"" href=""" & HTTPS_CURRENTURL & "login_menu.asp"">My Page</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""求職者の検索"" href=""" & HTTP_CURRENTURL & "company/myorderlist.asp"">求職者の検索</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""求職者検索条件管理"" href=""" & HTTP_CURRENTURL & "company/searchstaffcondition/list.asp"">求職者検索条件管理</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""ウォッチリスト"" href=""" & HTTP_CURRENTURL & "company/watchlist.asp"">ウォッチリスト</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""気になリスト"" href=""" & HTTP_CURRENTURL & "company/report/footprint.asp"">気になリスト</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""メール履歴"" href=""" & HTTPS_CURRENTURL & "company/mailhistory_company.asp"">メール履歴</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""一括メール管理"" href=""" & HTTPS_CURRENTURL & "company/lumpmail/list.asp"">一括メール管理</a></li>"
			Response.Write "<li class=""sidemenu_company""><a title=""求人票の修正"" href=""" & HTTP_CURRENTURL & "company/myorderlist.asp"">求人票の修正</a></li>"

			If Session("usertype") = "company" Then
				Response.Write "<li class=""sidemenu_company""><a href=""" & HTTPS_CURRENTURL & "company/company_reg1.asp"">自社情報を更新</a></li>"
				If G_IMAGELIMIT > 0 Then
					Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "company/img_upload.asp"">企業写真画像掲載</a></li>"
				End If

				If G_IMAGELIMIT > 1 Then
					Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "company/company_img_list.asp"">求人票用画像ストック</a></li>"
				End If
			ElseIf Session("usertype") = "dispatch" Then
				Response.Write "<li class=""sidemenu_company""><a href=""" & HTTPS_CURRENTURL & "dispatch/company_reg1.asp"">自社情報を更新</a></li>"
				If G_IMAGELIMIT > 0 Then
					Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "company/img_upload.asp"">企業写真画像掲載</a></li>"
				End If

				If G_IMAGELIMIT > 1 Then
					Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "company/company_img_list.asp"">求人票用画像ストック</a></li>"
				End If
			End If

			Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "mailtemplate/manager.asp"">メールテンプレート管理</a></li>"
			'Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "license/license_manager.asp"">ライセンス管理</a></li>
			If G_PLANTYPE = "mail" Then
				Response.Write "<li class=""sidemenu_company""><a href=""" & HTTP_CURRENTURL & "company/point/"">ポイント管理</a></li>"
			End If

			If G_PLANTYPE <> "mail" then
				Response.Write "<li class=""sidemenu_company""><a href=""" & HTTPS_CURRENTURL & "company/costperformance/"">採用改善ｻﾎﾟｰﾄｼｽﾃﾑ<img src=""/img/new.gif"" border=""0""></a></li>"
			End If

			Response.Write "<li class=""sidemenu_company_end""><a href=""" & HTTPS_CURRENTURL & "company/changepassword.asp"">パスワード変更</a></li>"
			Response.Write "<li class=""sidemenu_company_bottom""></li>"
		Else

			Response.Write "<ul>"
			Response.Write "<li class=""sidemenu_company_big"">ログイン</li>"
			Response.Write "<li>"
			Response.Write "<form id=""frmlogin"" method=""post"" action=""" & HTTPS_CURRENTURL & "login_check.asp"">"
			Response.Write "<div style=""line-height:22px; color:#6666cc; font-size:11px; border-right:solid 1px #9999ff; border-left:solid 1px #9999ff; padding-right:3px;"" align=""right"">"

			If G_SSLFLAG = False Then
				Response.Write "<a href=""" & G_URLS & """ style=""color:#0045f9;""><img src=""/img/common/security_key.gif"" border=""0"" height=""12"" alt="""">ＳＳＬをＯＮにする (推奨)</a><br>"
			Else
				Response.Write "<a href=""" & G_URL & """ style=""color:#0045f9;"">ＳＳＬをＯＦＦにする</a><br>"
			End If

			Response.Write "<font size=""1"" style=""width:50px; text-align:right; font-weight:bold;"">I　D</font>"
			Response.Write "<input type=""text"" name=""CONF_UserID"" value=""" & Request.Cookies("id_memory") & """ style=""width:100px;""><br>"
			Response.Write "<font size=""1"" style=""width:50px; text-align:right; font-weight:bold;"">パスワード</font>"
			Response.Write "<input type=""password"" name=""CONF_Password"" size=""11"" value="""" style=""width:100px;""><br>"
			Response.Write "<div style=""text-align:right;"">"

			If Request.QueryString("JUMP_URL_FLAG") = "True" Then
				For Each name In Request.QueryString
					Response.Write "<input type=""hidden"" name=""" & name & """ value=""" & Request.QueryString(name) & """>"
				Next
			End If

			Response.Write "<label><input type=""checkbox"" name=""frmautologinflag"" value=""1"">自動ﾛｸﾞｲﾝ</label>[<span style=""color:#0045f9; cursor:pointer;"" onclick=""window.open('/infomation/autologin.asp','autologin','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=400,height=220');""><u>？</u></span>]"
			Response.Write "<input type=""Submit"" name=""Login"" value=""ログイン"" onclick=""LoginCheckIdreg(); return false"" style=""font-size:12px; margin-right:1px;""><br>"
			Response.Write "</div>"
			Response.Write "</div>"
			Response.Write "<script type=""text/javascript""><!-- document.forms[0].UserID.focus(); // --></script>"
%><!-- #INCLUDE FILE="../error/errhandle.asp" --><%
			Response.Write "</form>"
			Response.Write "</li>"
			Response.Write "<li class=""sidemenu_company_bottom""></li>"
		End If

		Response.Write "<li class=""sidemenu_big"">求人広告</li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/index.asp"" title=""しごとナビ求人広告とは"">しごとナビ求人広告</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_function.asp"" title=""サービス概要"">サービス概要</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_voice.asp"" title=""ご利用企業様の声"">ご利用企業様の声</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/research.asp"" title=""人材採用方法診断"">人材採用方法診断</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_staffdata.asp"" title=""人材Ｄａｔａ"">人材Ｄａｔａ</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "jinzaisearch/index.asp"" title=""人材お試し検索"">人材お試し検索</a></li> "
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/charge.asp"" title=""料金システム"">料金システム</a></li>"

		If G_USERTYPE = "" Then
			Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/costperformance/"" title=""採用改善ｻﾎﾟｰﾄｼｽﾃﾑ"">採用改善ｻﾎﾟｰﾄｼｽﾃﾑ<img src=""/img/new.gif"" border=""0""></a></li>"
		End If

		'Response.Write "<li class=""sidemenu""><a href=""http://jinzai.shigotonavi.co.jp/joboffer/make_advertisement.asp"" target=""blank_"" title=""求人広告作成について"">求人広告作成について</a></li>"
		Response.Write "<li class=""sidemenu_end""><a href=""" & HTTPS_CURRENTURL & "company/request01.asp"" title=""お申し込み"">お申し込み</a></li>"
		Response.Write "<li class=""sidemenu_bottom""></li>"

		Response.Write "<li class=""sidemenu_big"">人材サービス</li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_introduce.asp"" title=""人材紹介"">人材紹介</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_temptoperm.asp"" title=""紹介予定派遣"">紹介予定派遣</a></li>"
		Response.Write "<li class=""sidemenu_end""><a href=""" & HTTP_CURRENTURL & "company/c_dispatch.asp"" title=""人材派遣"">人材派遣</a></li>"
		Response.Write "<li class=""sidemenu_bottom""></li>"

		'TOP 08/05/20 Lis林 ＦＣ ADD → 08/09/04 Lis林 DEL
		'Response.Write "<ul class=""sidemenulink"">"
		'Response.Write "<li>&nbsp;<a href=""" & HTTP_CURRENTURL & "company/fc_index.asp"" title=""人材サービスフランチャイズ:しごとナビFC"">しごとナビFC</a></li>"
		'Response.Write "</ul><br>"
		'BTM 08/05/20 Lis林 ＦＣ ADD → 08/09/04 Lis林 DEL

		Response.Write "<li class=""sidemenu_big"">サポート</li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_successpoint.asp"" title=""しごとナビ活用ブック"">しごとナビ活用ブック</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_scout3point.asp"">スカウトメール作成のコツ</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/qa.asp"">Ｑ＆Ａ</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_kiyaku.asp"">利用規約</a></li>"
		Response.Write "<li class=""sidemenu_end""><a href=""" & HTTPS_CURRENTURL & "company/access.asp"">お問合せ</a></li>"
		Response.Write "<li class=""sidemenu_bottom""></li>"

		Response.Write "<li class=""sidemenu_big"">会社概要</li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "lis/lis_annai.asp"" title=""会社案内"">会社案内</a></li>"
		Response.Write "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "lis/service-development.asp"" title=""総合人材サービス展開"">総合人材サービス展開</a></li>"
		Response.Write "<li class=""sidemenu_end""><a href=""" & HTTP_CURRENTURL & "lis/lis_saiyou.asp"" title=""採用情報"">採用情報</a></li>"
		Response.Write "<li class=""sidemenu_bottom""></li>"
		Response.Write "</ul>"

		Response.Write "<div align=""center"" style=""width:100%;padding-top:20px;"">"
		Response.Write "<a href=""/lis/blog_kimura.asp"">"
		Response.Write "<img src=""/img/top/top_blogBanner.gif"" border=""0"" alt=""木村亮郎のヒトビジネスつれづれ"">"
		Response.Write "</a>"
		Response.Write "</div>"

		Response.Write "<div align=""center"" style=""margin-top:10px;"">"
		Response.Write "<a href=""http://privacymark.jp/"" target=""_blank""><img src=""/img/privacy/p_75.gif"" alt=""プライバシーマーク"" border=""0""></a><br>"
		Response.Write "<a href=""" & HTTP_CURRENTURL & "privacy/privacy.asp"">個人情報保護について</a>"
		Response.Write "</div>"
		Response.Write "<div style=""text-align:center""></div>"

		Response.Write "</div>"
	ElseIf SidemenuType = 3 Then '共用
		If Session("usertype") = "staff" Then '求職者ログインしている場合
			Call NaviSidemenu(1)
		ElseIf Session("usertype") = "company" Or Session("usertype") = "dispatch" Then '企業ログインしている場合
			Call NaviSidemenu(2)
		Else
			Response.Write "</div>" 'メインコンテンツの幅指定divの閉め（開始はheader最下部）
			Response.Write "<div id=""idNavigation"" style=""width: 170px; float: left;"">"
			Response.Write "<a href=""" & HTTPS_CURRENTURL & "staff/person_reg1.asp""><img src=""/img/common/reg1_button.jpg"" alt=""しごとナビ会員登録"" border=""0"" style=""margin:3px 0px 2px 0px;""></a><br>"
			Response.Write "<div align=""right"" style=""font-size:11px; margin-bottom:5px;"">"
			Response.Write "<a href=""" & HTTPS_CURRENTURL & "login_menu.asp"">会員登録がお済みの方はこちら</a>"
			Response.Write "</div>"

			Response.Write "<ul>"
			Response.Write "<li class=""sidemenu_big"">お仕事をお探しの方</li>"
			Response.Write "<li class=""sidemenu""><a title=""お仕事検索"" href=""" & HTTP_CURRENTURL & "order/order_search_detail.asp"">お仕事検索</a></li>"
			Response.Write "<li class=""sidemenu""><a title=""ご利用ガイド"" href=""" & HTTP_CURRENTURL & "staff/s_aboutnavi.asp"">ご利用ガイド</a></li>"
			Response.Write "<li class=""sidemenu""><a title=""Ｑ＆Ａ"" href=""" & HTTP_CURRENTURL & "staff/qa.asp"">Ｑ＆Ａ</a></li>"
			Response.Write "<li class=""sidemenu""><a title=""利用規約"" href=""" & HTTP_CURRENTURL & "staff/s_kiyaku.asp"">利用規約</a></li>"
			Response.Write "<li class=""sidemenu_end""><a title=""お問合せ(求職者専用)"" href=""" & HTTPS_CURRENTURL & "staff/access.asp"">お問合せ(求職者専用)</a></li>"
			Response.Write "<li class=""sidemenu_bottom""></li>"
			Response.Write "</ul>"

			Response.Write "<ul>"
			Response.Write "<li class=""sidemenu_big"">人材をお探しの企業様</li>"
			Response.Write "<li class=""sidemenu""><a title=""ログイン"" href=""" & HTTPS_CURRENTURL & "login_menu.asp"">ログイン</a></li>"
			Response.Write "<li class=""sidemenu""><a title=""求人広告について"" href=""" & HTTP_CURRENTURL & "company/c_hajime.asp"">求人広告について</a></li>"
			Response.Write "<li class=""sidemenu""><a title=""人材紹介について"" href=""" & HTTP_CURRENTURL & "company/c_introduce.asp"">人材紹介について</a></li>"
			Response.Write "<li class=""sidemenu""><a title=""紹介予定派遣について"" href=""" & HTTP_CURRENTURL & "company/c_temptoperm.asp"">紹介予定派遣について</a></li>"
			Response.Write "<li class=""sidemenu""><a title=""人材派遣について"" href=""" & HTTP_CURRENTURL & "company/c_dispatch.asp"">人材派遣について</a></li>"
			Response.Write "<li class=""sidemenu""><a title=""お問合せ(求人企業様専用)"" href=""" & HTTPS_CURRENTURL & "company/access.asp"">お問合せ(求人企業様専用)</a></li>"
			Response.Write "<li class=""sidemenu_end""><a title=""会社案内"" href=""" & HTTP_CURRENTURL & "lis/lis_annai.asp"">会社案内</a></li>"
			Response.Write "<li class=""sidemenu_bottom""></li>"
			Response.Write "</ul>"

			Response.Write "<!-- SIDE-MENU END -->"
			Response.Write "<br>"
			Response.Write "<center>"
			Response.Write "<a href=""http://privacymark.jp/"" target=""_blank""><img src=""/img/privacy/p_75.gif"" alt=""プライバシーマーク"" border=""0""></a><br>"
			Response.Write "<a href=""" & HTTP_CURRENTURL & "privacy/privacy.asp"">個人情報保護について</a>"
			Response.Write "</center>"

			Response.Write "</div>"
		End If
	End If
End Function

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
%><!-- #INCLUDE VIRTUAL="/include/ads/navifooter.asp" --><%
	Response.Write "<br>"
	Response.Write "<div style=""text-align:left;height:55px; width:785px;"">"
	Response.Write "<ul class=""footer"" style=""float:left;padding-left:5px;"">"
	Response.Write "<li style=""float:left;""><a href=""" & HTTP_CURRENTURL & """ title=""転職・求人サイトしごとナビ"" class=""topdecnone"">ＨＯＭＥ</a></li>"
	Response.Write "<li style=""float:left;"">｜<a href=""" & HTTP_CURRENTURL & "staff/Ranking.asp"" title=""求職者ランキング"" class=""topdecnone"">求職者ランキング</a></li>"
	'Response.Write "<li style=""float:left;"">｜<a href=""/company/c_hajime.asp"" class=""topdecnone"">求人広告</a></li>"
	Response.Write "<li style=""float:left;"">｜<a href=""" & HTTP_CURRENTURL & "infomation/info.asp"" title=""記事倉庫"" class=""topdecnone"">記事倉庫</a></li>"
	Response.Write "<li style=""float:left;"">｜<a href=""" & HTTP_CURRENTURL & "lis/lis.asp"" title=""運営会社・当社採用情報"" class=""topdecnone"">運営会社・当社採用情報</a></li>"
	'Response.Write "<li style=""float:left;"">｜<a href=""/staff/s_aboutnavi.asp"" title=""ご利用ガイド"" class=""topdecnone"">ご利用ガイド</a></li>"
	'Response.Write "<li style=""float:left;"">｜<a href=""/staff/qa.asp"" title=""Ｑ＆Ａ"" class=""topdecnone"">Ｑ＆Ａ</a></li>"
	'Response.Write "<li style=""float:left;"">｜<a href=""/staff/s_kiyaku.asp"" title=""利用規約"" class=""topdecnone"">利用規約</a></li>"
	Response.Write "<li style=""float:left;"">｜<a href=""" & HTTPS_CURRENTURL & "staff/access.asp"" title=""お問合せ"" class=""topdecnone"">お問合せ&lt;転職希望の方向け&gt;</a></li>"
	Response.Write "<li style=""float:left;"">｜<a href=""" & HTTP_CURRENTURL & "s_contents/s_books.asp"" title=""転職に役立つ本"" style=""margin-left:5px;"" class=""topdecnone"">転職に役立つ本</a></li>"
	'Response.Write "<li style=""float:left;"">｜<a href=""/link.asp"" title=""リンクポリシー"" class=""topdecnone"">リンクポリシー</a></li>"
	'Response.Write "<li style=""float:left;"">｜<a href=""/link_collection.asp"" title=""お役立ち厳選リンク集"" class=""topdecnone"">お役立ち厳選リンク集</a></li>"
	Response.Write "<li style=""float:left;"">｜<a href=""" & HTTP_CURRENTURL & "shigotonavi/sitemap.asp"" class=""topdecnone"" title=""サイトマップ"">サイトマップ</a></li>"
	Response.Write "</ul>"
	Response.Write "<br clear=""all"">"

	Response.Write "<div style=""width:100%; height:5px; margin:0px; padding:0px; background-image:url(/img/footer/footer_1.gif); background-repeat:repeat-x;"">"
	Response.Write "</div>"
	Response.Write "<div style=""float:left;width:580px;padding-left:5px;"">"
	Response.Write "<ul class=""footer"">"

	Response.Write "<li style=""float:left;""><a href=""" & HTTP_CURRENTURL & "company/index.asp"" title=""企業向けコンテンツ"" class=""topdecnone"">企業向けコンテンツ</a></li>"
	Response.Write "<li style=""float:left;"">｜<a href=""" & HTTP_CURRENTURL & "company/c_hajime.asp"" title=""求人広告"" class=""topdecnone"">求人広告</a></li>"
	Response.Write "<li style=""float:left;"">｜<a href=""" & HTTP_CURRENTURL & "company/c_dispatch.asp"" title=""人材派遣"" class=""topdecnone"">人材派遣</a></li>"
	Response.Write "<li style=""float:left;"">｜<a href=""" & HTTP_CURRENTURL & "company/c_introduce.asp"" title=""人材紹介"" class=""topdecnone"">人材紹介</a></li>"
	Response.Write "<li style=""float:left;"">｜<a href=""" & HTTP_CURRENTURL & "company/c_temptoperm.asp"" title=""紹介予定派遣"" class=""topdecnone"">紹介予定派遣</a>｜</li>"
	'Response.Write "<li style="float:left;">｜<a href=""" & HTTPS_CURRENTURL & "company/fc_index.asp" title="人材サービスフランチャイズ,しごとナビFC" class="topdecnone">しごとナビFC</a>｜</li>"
	Response.Write "</ul><br>"
	Response.Write "<ul class=""footer"">"
	Response.Write "<li style=""float:left;""><a href=""" & HTTPS_CURRENTURL & "company/access.asp"" title=""お問合せ&lt;企業様向け&gt;"" class=""topdecnone"">お問合せ&lt;企業様向け&gt;</a></li>"
	Response.Write "<li style=""float:left;"">｜<a href=""" & HTTP_CURRENTURL & "company/c_staffdata.asp"" title=""しごとナビ求職者と掲載企業データ"" class=""topdecnone"">しごとナビ求職者と掲載企業データ</a></li>"
	Response.Write "<li style=""float:left;"">｜<a href=""" & HTTPS_CURRENTURL & "company/access.asp"" title=""お問合せ&lt;企業様向け&gt;"" class=""topdecnone"">広告代理店の方のお問合せ</a></li>"
	Response.Write "</ul>"
	Response.Write "</div>"
	Response.Write "<div style=""float:right;width:180px;"">"
	Response.Write "<a href=""" & HTTP_LIS_CURRENTURL & """ title=""転職サイトしごとナビの運営会社-リス株式会社-"" target=""_blank""><img src=""/img/footer/footer_lis_logo_1.gif"" alt=""転職サイト｢しごとナビ｣運営-リス株式会社-"" border=""0""></a>"
	Response.Write "</div>"
	Response.Write "<br clear=""all"">"
	Response.Write "</div>"
	Response.Write "</div>"
	Response.Write "</div>" & vbCrLf

	'ページ全体の幅閉め（開始はheader最上部）
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
	Response.Write "<a href=""" & HTTP_CURRENTURL & "cafe/cafe_list.asp"" title=""ナビカフェ""><img src=""/img/rightmenu/navicafe_banner_top.jpg"" alt=""ナビカフェ"" border=""0"" style=""margin:0px;padding:0px;""></a>"
	Response.Write "<div style=""margin-top:0px;padding:14px 6px 0px 8px;font-size:10px;line-height:15px;"">"

	'** TOP 08/11/05 Lis林 ADD
	'現在掲載中＆TOP3のトピ
	sSQLnsr = "up_GetData_NC_Topic '','','','1','3'"
	flgQEnsr = QUERYEXE(dbconn, oRSnsr, sSQLnsr, sErrornsr)
	Do While GetRSState(oRSnsr) = True
		Response.Write "<a href='" & HTTP_CURRENTURL & "cafe/cafe_detail.asp?t=" & oRSnsr.Collect("TopicID")
		Response.Write "' title='" & oRSnsr.Collect("Title") & "'>・"
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
		Response.Write "<li class=""rightmenu""><a title=""ご利用ガイド"" href=""" & HTTP_CURRENTURL & "staff/s_aboutnavi.asp"">ご利用ガイド</a></li>"
		Response.Write "<li class=""rightmenu""><a title=""Ｑ＆Ａ"" href=""" & HTTP_CURRENTURL & "staff/qa.asp"">Ｑ＆Ａ</a></li>"
		Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "staff/s_searchexplanation.asp"" title=""お仕事検索方法"">お仕事検索方法</a></li>"
		Response.Write "<li class=""rightmenu""><a title=""利用規約"" href=""" & HTTP_CURRENTURL & "staff/s_kiyaku.asp"">利用規約</a></li>"
		Response.Write "<li class=""rightmenu_end""><a title=""お問合せ(求職者専用)"" href=""" & HTTPS_CURRENTURL & "staff/access.asp"">お問合せ(求職者専用)</a></li>"
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
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_mensetsu_index.asp"" title=""面接対策"">面接対策</a></li>"
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "column/column_1.asp"" title=""派遣社員-成功の鍵はプロ意識"">派遣社員<span style=""font-size:10px;"">-成功の鍵はプロ意識</span></a></li>"
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_kyuuyomeisai.asp"" title=""あなたの給与明細"">あなたの給与明細</a></li>"
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_ready.asp"" title=""転職の心構え"">転職の心構え</a></li>"
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_proce.asp"" title=""転職に必要な手続き"">転職に必要な手続き</a></li>"
	Response.Write "<li class=""rightmenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_goukaku.asp"" title=""合格率ＵＰマニュアル"">合格率ＵＰマニュアル</a></li>"
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

	If Session("usertype") = "staff" Then '求職者ログインしている場合	
	ElseIf Session("usertype") = "company" Or Session("usertype") = "dispatch" Then '企業ログインしている場合
	Else
	End If
End Function
%>
