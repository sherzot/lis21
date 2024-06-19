                                                                                                                                                                                                                                                                    <%
'******************************************************************************
'概　要：サイドメニュー
'引　数：SidemenuType	0【トップ】1【求職者】2【企業】3【共用】4【代理店】
'備　考：
'使用元：
'履　歴：2008/02/07 LIS K.Niina 作成
'　　　：2008/05/20 LIS M.Hayashi しごとナビFC追加
'　　　：2011/02/16 LIS K.Kokubo スパム的なtitle属性削除,しごとナビツイッターバナー削除
'      ：2015/11/20 LIS K.Kimura サイドメニューを変更
'******************************************************************************
Function NaviSidemenu(SidemenuType)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim iTabIndexType
	Dim sHTML10	'求職者ログイン部分
	Dim sHTML11	'企業ログイン部分
	Dim sHTML12	'求職者,企業My Menu
	Dim sHTML13	'アクセス課金
	Dim sHTML20	'TOPページのバナー
	Dim sHTML21	'求職者情報
	Dim sHTML30	'社長ブログバナー
	Dim sHTML31	'Pマーク
	Dim sHTML40	'社長ナビバナー(求職者向け)
	Dim sHTML41	'社長ナビバナー(企業向け)
	Dim sHTML60	'ツイッター
	Dim sHTML61	'人材紹介ツイッター
	Dim sHTML62	'東北地方太平洋沖地震の影響について
	Dim sHTML63	'スマホ
	Dim sHTML64	'Facebookページバナー
	Dim sHTML80	'リス社員募集
	Dim sHTML90	'派遣協会派遣スタッフアンケートバナー
	Dim sHTML91	'社内案件急募バナー
	Dim sHTML100 '学ぶのヤツ
	Dim sHTML	'タブIndex毎のナビゲーション部分

	Dim sScript	'ログインチェックスクリプト

	Dim sParamName
	Dim si
	Dim sMidoku
	Dim rank(2)
	Dim rankcount(2)
	Dim idx
	Dim iCollectionCount
	
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

		'<ログインチェックスクリプト>
		sScript = ""
		sScript = sScript & "<script type=""text/javascript"">"
		sScript = sScript & "function LoginCheckIdreg(){"
		sScript = sScript & "var ofrm = document.forms.frmlogin;"
		sScript = sScript & "if(!navigator.cookieEnabled) {"
		sScript = sScript & "alert('cookie（クッキー）の利用ができない設定になっています。\nブラウザやセキュリティーソフトのcookie設定をご確認下さい。');"
		sScript = sScript & "return false;"
		sScript = sScript & "}"
		sScript = sScript & "if(ofrm.CONF_UserID.value.length === 0){alert('認証IDを入力してください。');return false;}"
		sScript = sScript & "if(ofrm.CONF_Password.value.length === 0){alert('パスワードを入力してください。');return false;}"
		sScript = sScript & "if(ofrm.CONF_Password.value.length < 3 || ofrm.CONF_Password.value.length > 20){alert('パスワードは３～２０文字で入力してください。');return false;}"
		sScript = sScript & "ofrm.submit();"
		sScript = sScript & "}"
		sScript = sScript & "</script>"
		'</ログインチェックスクリプト>

        '<動画バナー>
'            sHTML10 = sHTML10 & "<a href=""https://www.youtube.com/watch?v=T3n06VU8T-Q" & HTTPS_CURRENTURL & "valueoffer/""TARGET="_blank"><img src=""/img/common/tutrial_banner01.png"" alt=""バリューオファー"" style=""margin-top:3px;""></a>"
            sHTML10 = sHTML10 & "<a href=""https://www.youtube.com/watch?v=T3n06VU8T-Q" & HTTPS_CURRENTURL & "valueoffer/""target=""_blank""><img src=""/img/common/tutrial_banner01.png"" alt=""バリューオファー"" style=""margin-top:3px;""></a>"
        'ポイントＰＲバナー張替
            'sHTML10 = sHTML10 & "<a href=""https://youtu.be/9FlpxFA6TYc""target=""_blank""><img src=""/img/common/conpri_banner1.png"" alt=""バリューオファー"" style=""margin-top:3px;""></a>"
        '</動画バナー>
        sHTML10 = sHTML10 & "<a href=""" & HTTPS_CURRENTURL & "pr/pushpoint.asp""><img src=""/img/common/how_to_use2.png""></a>"

	'以下の部分からスクロールに追従するサイドメニュー
	sHTML10 = sHTML10 & "<div class=""floatingmenu"" id=""moveside""><div align=""center"">"
        if GetForm("ordercode", 2) <> "" then
			if IsRE(Trim(Replace(Server.HTMLEncode(GetForm("ordercode", 2)), "'", "’")), "^J\d\d\d\d\d\d\d$", True) = True then
				sHTML10 = sHTML10 & "<a href=""" & HTTPS_CURRENTURL & "staff/person_reg1.asp?ordercode=" & GetForm("ordercode", 2) & """><img src=""/img/common/reg1_button_big_3.png"" border=""0"" alt=""会員登録(無料)"" style=""margin-top:3px;""></a>"
			else
				sHTML10 = sHTML10 & "<a href=""" & HTTPS_CURRENTURL & "staff/person_reg1.asp""><img src=""/img/common/reg1_button_big_3.png"" border=""0"" alt=""会員登録(無料)"" style=""margin-top:3px;""></a>"
			end if
		else
			sHTML10 = sHTML10 & "<a href=""" & HTTPS_CURRENTURL & "staff/person_reg1.asp""><img src=""/img/common/reg1_button_big_3.png"" border=""0"" alt=""会員登録(無料)"" style=""margin-top:3px;""></a>"
		end if

		sHTML10 = sHTML10 & "<a href=""" & HTTPS_CURRENTURL & "/point/pr/""target=""_blank""><img src=""/img/neo/point_present.png""></a>"

		sHTML10 = sHTML10 & "<script type=""text/javascript""><!-- document.forms[0].UserID.focus(); // --></script>"
		sHTML10 = sHTML10 & "</div>"

		'<コンサル紹介>
		'2016/04/11 木村追加
		'2016/04/21 3人しかいないので非表示
		'sHTML10 = sHTML10 & "<a href=""" & HTTPS_CURRENTURL & "consultant/consultantbranch.asp""><img src=""/img/common/con_int.png"" alt=""在籍コンサルタント紹介"" style=""margin-top:3px;border:1px solid #000;""></a>"
		'</コンサル紹介>

        '<バリューオファー物語>
        'If Request.ServerVariables("PATH_INFO") = "/staff/s_resume_kakikata.asp" Then
            'sHTML10 = sHTML10 & "<div align=""center"" style=""margin: 9px 0px 5px;"">"
            'sHTML10 = sHTML10 & "<a href=""" & HTTPS_CURRENTURL & "valueoffer/persona.asp""><div><img src=""/img/C_K_NAVI.GIF"" height=""50"">バリューオファー物語</div></a>"
            'sHTML10 = sHTML10 & "</div>"
            'sHTML10 = sHTML10 & "<a href=""" & HTTPS_CURRENTURL & "valueoffer/persona02.asp""><img src=""/img/common/persona_banner02.png"" alt=""バリューオファー"" style=""margin-top:3px;""></a>"

        'End If
        '</バリューオファー物語>

		
		'sHTML10 = sHTML10 & "<form id=""mailReg""><input type=""text"" value=""mail""><br><input type=""button"" value=""メルマガ登録"" onClick=""location.href='/staff/mailReg.asp'""></form>"
		'sHTML10 = sHTML10 & "<a href=""http://www.shigotonavi.co.jp/iphone/index.html"" target=""_blank""><img src=""/img/link/iphone_banner.png"" style=""width: 210px;""></a>"
		'sHTML10 = sHTML10 & "<a href=""http://www.a-rirekisyo.jp/"" target=""_blank""><img src=""/img/link/a-resume_banner.gif"" style=""width: 210px;""></a>"
		'sHTML10 = sHTML10 & "<a href=""/recruit/se/""><img src=""/recruit/img/banner.png""></a>"

		'<企業ログインフォーム>

		'</企業ログインフォーム>
	End If

	'<求職者My Menu>
    '2015/08/19 木村改修：段組みメニュー
	If G_USERTYPE = "staff" Then
		'未読メール件数
		sSQL = "SELECT COUNT(*) AS Cnt FROM MailHistory WITH(NOLOCK) WHERE ReceiverCode ='" & G_USERID & "' AND OpenDay IS NULL AND ReceiverDelFlag = '0'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			If oRS.Collect("Cnt") = 0 Then
				sMidoku = "(<img src=""/img/staff/mail/mailhei.gif"" border=""0"" alt="""" style=""margin:0px 1px;"">未読" & oRS.Collect("Cnt") & "件)"
			Else
				sMidoku = "(<span style=""color:#ff0000; font-weight:bold;""><img src=""/img/staff/mail/mailhei.gif"" border=""0"" alt="""" style=""margin:0px 1px;"">未読" & oRS.Collect("Cnt") & "件</span>)"
			End If
		End If

		sHTML12 = sHTML12 & "<div id=""moveside""><ul class=""smartSidenone"">"

         '<動画バナー>
'            sHTML10 = sHTML10 & "<a href=""https://www.youtube.com/watch?v=T3n06VU8T-Q" & HTTPS_CURRENTURL & "valueoffer/""><img src=""/img/common/tutrial_banner01.png"" alt=""バリューオファー"" style=""margin-top:3px;""></a>"
            sHTML10 = sHTML10 & "<a href=""https://www.youtube.com/watch?v=T3n06VU8T-Q" & HTTPS_CURRENTURL & "valueoffer/""target=""_blank""><img src=""/img/common/tutrial_banner01.png"" alt=""バリューオファー"" style=""margin-top:3px;""></a>"
        'ポイントＰＲバナー張替
             sHTML10 = sHTML10 & "<a href=""" & HTTPS_CURRENTURL & "/point/pr""target=""_blank""><img src=""/img/neo/point_present.png""style=""margin-top:3px;""></a>"
        '</動画バナー>


        '2015/08/28 なし
        'If G_SSLFLAG = False Then
		'sHTML12 = sHTML12 & "<li class=""sidemenu_staff_big"">My&nbsp;メニュー&nbsp;(<a href=""" & HTTP_CURRENTURL & "logout.asp"">ログアウト</a>)</li>"
        'Else
		sHTML12 = sHTML12 & "<li class=""sidemenu_staff_big"">My&nbsp;メニュー&nbsp;(<a href=""" & HTTPS_CURRENTURL & "logout.asp"">ログアウト</a>)</li>"
        'End IF
        '共通
		sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/s_login.asp"">My&nbsp;ページ</a></li>"
        sHTML12 = sHTML12 & "<li class=""sidemenu"" style=""border-bottom:none;""><a class=""nobottom"" href=""" & HTTPS_CURRENTURL & "staff/person_detail.asp"">プロフィール管理</a></li>"

		'バリューオファー
		'2015/03/02 池田改修
        Dim sMikaitou
		'未読メール件数
		sSQL = "EXEC up_ExistsOfferStep2c '" & G_USERID & "'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			If oRS.RecordCount = 0 Then
				sMikaitou = "(<img src=""/img/staff/mail/mailhei.gif"" border=""0"" alt="""" style=""margin:0px 1px;"">未回答" & oRS.RecordCount & "件)"
			Else
				sMikaitou = "(<span style=""color:#ff0000; font-weight:bold;""><img src=""/img/staff/mail/mailhei.gif"" border=""0"" alt="""" style=""margin:0px 1px;"">未回答" & oRS.RecordCount & "件</span>)"
			End If
		End If
		Call RSClose(oRS)

        'ポイント申請
		'sHTML12 = sHTML12 & "<li class=""sidetitle"">GPoint申請<img src=""/img/c_new.gif"" alt="""" border=""0"">"
		'sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/apply_GPoint.asp?PointType=login"">ログインポイント（1日1回）</a></li>"
		'sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/apply_GPoint.asp?PointType=DRegist"">登録ポイント</a></li>"

		sHTML12 = sHTML12 & "<li class=""sidetitle"">バリューオファー</li>"
		sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/step2a.asp"">希望条件入力</a></li>"
		sHTML12 = sHTML12 & "<li class=""sidemenu""><a class=""nobottom"" href=""" & HTTPS_CURRENTURL & "staff/step2c.asp?offer_ques=true"">あなたに興味を持った<br>企業からの質問" & sMikaitou & "</a></li>"

        '応募・状況
        sHTML12 = sHTML12 & "<li class=""sidetitle"">転職サポート</li>"
        sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/my_footprint.asp"">閲覧履歴</a></li>"
        sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/watchlist.asp"">お気に入りリスト</a></li>"
		sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/edit_list.asp"">応募一覧</a></li>"
		sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/mailhistory_person.asp"">メール管理" & sMidoku & "</a></li>"
		sHTML12 = sHTML12 & "<li class=""sidemenu""><a class=""nobottom"" href=""" & HTTPS_CURRENTURL & "staff/schedule/"">スケジュール管理</a></li>"


        '2015/09/01　要改修、現状あまり意味がないページ
		'sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/jobcon/"">ジョブ・コンシェルジュ</a></li>"
        
        '<ステップ6対応>
		'sHTML12 = sHTML12 & "<li class=""sidetitle"">バリューオファー</li>"
		'sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/step2a.asp"">希望条件入力</a></li>"
		'sHTML12 = sHTML12 & "<li class=""sidemenu""><a class=""nobottom"" href=""" & HTTPS_CURRENTURL & "staff/step2c.asp"">企業からの質問" & sMikaitou & "</a></li>"
		'</ステップ6対応>

        '履歴書
        sHTML12 = sHTML12 & "<li class=""sidetitle"">履歴書・職務経歴書の印刷</li>"
		sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/resume_print.asp"">履歴書・職務経歴書印刷</a></li>"
		sHTML12 = sHTML12 & "<li class=""sidemenu""><a class=""nobottom"" href=""" & HTTPS_CURRENTURL & "staff/resume_picture.asp"">履歴書用写真登録</a></li>"
		'sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/resumemanual.pdf"" target=""_blank"">履歴書作成マニュアル</a></li>"
        '停止退会
        sHTML12 = sHTML12 & "<li class=""sidetitle"">各種設定</li>"
		sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/person_edit6.asp"">希望条件（メール配信条件）</a></li>"
		sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/notification_mail_service.asp"">スケジュール通知</a></li>"

        '2015/09/01　要改修、現状あまり意味がないページ
        'sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/searchordercondition/"">検索条件管理</a></li>"
		sHTML12 = sHTML12 & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/changepassword.asp"">パスワードの変更</a></li>"
        sHTML12 = sHTML12 & "<li class=""sidemenu_end""><a href=""" & HTTPS_CURRENTURL & "suspension/questionnarie.asp"">休止・退会</a></li>"

		sHTML12 = sHTML12 & "</ul></div><!--/#moveside-->"
'		sHTML12 = sHTML12 & "<a href=""/neo/oiwai/"" target=""_blank"" id=""oiwai_page"">ポイント申請</a>"
		sHTML12 = sHTML12 & "<a href=""/point/pr/"" target=""_blank""><img src=""/img/neo/point_present.png""></a>"
		'sHTML12 = sHTML12 & "<a href=""/recruit/se/""><img src=""/recruit/img/banner.png""></a>"
	End If
	'</求職者My Menu>

	'<企業My Menu>
	If G_USERTYPE = "company" Then
		'<求人数取得>
		iCollectionCount = 0
		sSQL = "EXEC sp_GetCollectionCount '" & G_USERID & "';"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			iCollectionCount = oRS.Collect("Cnt")
		End If
		Call RSClose(oRS)
		'</求人数取得>

		sHTML12 = sHTML12 & "<ul class=""smartSidenone"">"
		sHTML12 = sHTML12 & "<li class=""sidemenu_company_big"">My&nbsp;メニュー&nbsp;(<a href=""" & HTTP_CURRENTURL & "logout.asp"">ログアウト</a>)</li>"
		sHTML12 = sHTML12 & "<li class=""sidemenu_end""><a href=""" & HTTPS_CURRENTURL & "management/index.asp"" target=""_blank"">管理画面ページ</a></li>"
		sHTML12 = sHTML12 & "</ul>"

	End If
	'</企業My Menu>

	'<TOPページのナビゲーションバナー広告>
	If SideMenuType = 0 Then

		'<履歴書自動作成>
		sHTML20 = sHTML20 & "<ul>"
		sHTML20 = sHTML20 & "<li class=""sidemenu_big"">便利ツール</li>"
		sHTML20 = sHTML20 & "<li style=""border-left:solid 1px #cccccc; border-right:solid 1px #cccccc; line-height:17px;"">"
		sHTML20 = sHTML20 & "<a href=""/staff/s_resume.asp"" style=""display:block; background-image:url(/img/sidemenu/resume_banner.jpg); width:154px; height:73px; font-size:10px; padding:54px 0px 0px 14px; color:#444444; text-decoration:none;"">"
		sHTML20 = sHTML20 & "必要な項目を入力するだけで完成！<br>55万人が使う安心のサービス！<br>自分に合った履歴書が作れる！"
		sHTML20 = sHTML20 & "</a>"
		sHTML20 = sHTML20 & "</li>"
		sHTML20 = sHTML20 & "</ul>"


	End If
	'</TOPページのナビゲーションバナー広告>

	'<求職者情報>
	If iTabIndexType = 0 Then
		idx = 0
		sSQL = "SELECT TOP 3 Subitem,Number FROM Person_Statistics WHERE item = '都道府県別' ORDER BY CONVERT(INT,Number) DESC;"
		flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
		Do While GetRSState(oRS) = True
			rank(idx) = Replace(Replace(Replace(oRS.Collect("SubItem"),"都",""),"府",""),"県","")
			rankcount(idx) = oRS.Collect("Number")
			idx = idx + 1
			oRS.MoveNext
		Loop
		Call RSClose(oRS)

		sHTML21 = sHTML21 & "<div align=""center"" style=""width:100%;"">"
		sHTML21 = sHTML21 & "<div class=""sidemenu_big"" style=""text-align:left;"">求職者情報</div>"
		sHTML21 = sHTML21 & "<div style=""border-left:solid 1px #cccccc; border-right:solid 1px #cccccc; background-image:url(/img/sidemenu/jinzaidata_background.gif);"" align=""center"">"
		sHTML21 = sHTML21 & "<table style=""width:155px; font-size:10px; text-align:left;"">"
		sHTML21 = sHTML21 & "<tr>"
		sHTML21 = sHTML21 & "<td>都道府県別</td>"
		sHTML21 = sHTML21 & "<td>1位:" & rank(0) & "</td>"
		sHTML21 = sHTML21 & "<td align=""right"">" & rankcount(0) & "名</td>"
		sHTML21 = sHTML21 & "</tr>"
		sHTML21 = sHTML21 & "<tr>"
		sHTML21 = sHTML21 & "<td></td>"
		sHTML21 = sHTML21 & "<td>2位:" & rank(1) & "</td>"
		sHTML21 = sHTML21 & "<td align=""right"">" & rankcount(1) & "名</td>"
		sHTML21 = sHTML21 & "</tr>"
		sHTML21 = sHTML21 & "<tr>"
		sHTML21 = sHTML21 & "<td></td>"
		sHTML21 = sHTML21 & "<td>3位:" & rank(2) & "</td>"
		sHTML21 = sHTML21 & "<td align=""right"">" & rankcount(2) & "名</td>"
		sHTML21 = sHTML21 & "</tr>"

		idx = 0
		sSQL = "SELECT TOP 3 item,subitem, Number FROM Person_Statistics where item = '10歳代' or item = '20歳代' or item = '30歳代' or item = '40歳代' or item = '50歳代' or item = '60歳以上' order by convert(int,Number) desc"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		Do While GetRSState(oRS) = True
			rank(idx) = Replace(oRS.Fields("Item").Value,"歳","") & oRS.Fields("SubItem").Value
			rankcount(idx) = oRS.Fields("Number").Value
			idx = idx + 1
			oRS.MoveNext
		Loop
		Call RSClose(oRS)

		sHTML21 = sHTML21 & "<tr>"
		sHTML21 = sHTML21 & "<td>年齢別</td>"
		sHTML21 = sHTML21 & "<td>1位:" & rank(0) & "</td>"
		sHTML21 = sHTML21 & "<td align=""right"">" & rankcount(0) & "名</td>"
		sHTML21 = sHTML21 & "</tr>"
		sHTML21 = sHTML21 & "<tr>"
		sHTML21 = sHTML21 & "<td></td>"
		sHTML21 = sHTML21 & "<td>2位:" & rank(1) & "</td>"
		sHTML21 = sHTML21 & "<td align=""right"">" & rankcount(1) & "名</td>"
		sHTML21 = sHTML21 & "</tr>"
		sHTML21 = sHTML21 & "<tr>"
		sHTML21 = sHTML21 & "<td></td>"
		sHTML21 = sHTML21 & "<td>3位:" & rank(2) & "</td>"
		sHTML21 = sHTML21 & "<td align=""right"">" & rankcount(2) & "名</td>"
		sHTML21 = sHTML21 & "</tr>"
		sHTML21 = sHTML21 & "<tr>"
		sHTML21 = sHTML21 & "<td colspan=""3"" align=""right""><a href=""/company/c_staffdata.asp""><img src=""/img/sidemenu/kuwashiku_min.jpg"" alt=""詳しくはこちら"" border=""0""></a>"
		sHTML21 = sHTML21 & "</tr>"
		sHTML21 = sHTML21 & "</table>"
		sHTML21 = sHTML21 & "</div>"
		sHTML21 = sHTML21 & "<br style=""clear:both;"">"
		sHTML21 = sHTML21 & "</div>"
	End If
	'</求職者情報>

	'<Pマーク>
	sHTML31 = sHTML31 & "<div align=""center"" style=""margin:10px 0 5px 0;"">"
	sHTML31 = sHTML31 & "<a href=""http://privacymark.jp/"" target=""_blank""><img src=""/img/privacy/p_75.gif"" alt=""プライバシーマーク"" border=""0""></a><br>"
	sHTML31 = sHTML31 & "<a href=""" & HTTP_CURRENTURL & "privacy/privacy.asp"">個人情報保護について</a>"
	sHTML31 = sHTML31 & "</div>"
	sHTML31 = sHTML31 & "<div class=""center""></div>"
	'</Pマーク>



	'<社内案件急募バナー>
	If Date <= "2011/04/20" Then
		sHTML91 = sHTML91 & "<div style=""width:170px;margin-bottom:10px;"">"
		sHTML91 = sHTML91 & "<img src=""/img/banner/gu0001.jpg"" alt=""食品成分の分析経験者,勤務地は群馬県沼田市,急募"" border=""0"" style=""cursor:pointer;"" onclick=""location.href='/ad_banner_control/c_r.asp?origin=gu0001';"">"
		sHTML91 = sHTML91 & "</div>"
	End If
	'</社内案件急募バナー>

	If iTabIndexType = 0 Then
		'<はじめての方>
		sHTML = sHTML & "<ul>"
		sHTML = sHTML & "<li class=""sidemenu_big"">はじめての方</li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/person_reg1.asp"">会員登録（履歴書登録）</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "staff/qa.asp"">しごとナビQ&A</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "order/order_search_detail.asp"">求人情報検索</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/passwordreminder.asp"">ID・パスワードの再取得</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "promotion/mobilepromotion.asp"">しごとナビモバイル</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "promotion/conpri_riyou.asp"">コンビニプリント(セブン‐イレブン)の利用方法</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu_end""><a href=""" & HTTP_CURRENTURL & "promotion/s_conpri_riyou.asp"">履歴書印刷（ローソン・ファミリーマート・サークルK・サンクス）</a></li>"
        sHTML = sHTML & "<li class=""sidetitle"">履歴書・職務経歴書の印刷</li>"
        sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "promotion/conpripromotion.asp""><img style=""width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src=""/img/top/clogo_711.png""></a></li>"
        sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "promotion/s_conpri_riyou.asp""><img style=""width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src=""/img/top/clogo_familymart.png""></a></li>"
        sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "promotion/s_conpri_riyou.asp""><img style=""width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src=""/img/top/clogo_lawson.png""></a></li>"
		sHTML = sHTML & "</ul>"
        sHTML = sHTML & "</ul>"
		'</はじめての方>
	ElseIf iTabIndexType = 1 Then
		'<求人を探す>
		sHTML = sHTML & "<ul>"
		sHTML = sHTML & "<li class=""sidemenu_big"">求人を探す</li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "order/order_search_detail.asp"">求人情報を探す</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "order/order_list_accesscount.asp"">人気の求人情報トップ10</a></li>"
		sHTML = sHTML & "<li class=""sidemenu_end""><a href=""" & HTTP_CURRENTURL & "railway/railway_search1.asp"">沿線検索</a></li>"
		sHTML = sHTML & "</ul>"
		'</求人を探す>
	ElseIf iTabIndexType = 2 Then
		'<便利ツール>
		sHTML = sHTML & "<ul>"
		sHTML = sHTML & "<li class=""sidemenu_big"">便利ツール</li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "staff/s_resume.asp"">履歴書の自動作成</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "staff/s_resume_kakikata.asp"">履歴書の書き方</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "staff/s_resume_qa.asp"">履歴書Ｑ＆Ａ</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "staff/s_careersheet.asp"">職務経歴書の自動作成/フォーマットのダウンロード</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "staff/s_careersheet_kakikata_1.asp"">職務経歴書の書き方</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/motive_index.asp"">志望動機メーカー</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_jikopr.asp"">自己PRメーカー</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_taishokunegai.asp"">退職願の書き方</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_year_calculation.asp"">西暦・和暦/学歴計算</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "promotion/conpripromotion.asp"">履歴書印刷（セブン‐イレブン）</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "promotion/s_conpri_riyou.asp"">履歴書印刷（ローソン・ファミリーマート・サークルK・サンクス）</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu_end""><a href=""" & HTTP_CURRENTURL & "conpri/"">書類印刷サービス</a></li>"
        sHTML = sHTML & "<li class=""sidetitle"">履歴書・職務経歴書の印刷</li>"
        sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "promotion/conpripromotion.asp""><img style=""width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src=""/img/top/clogo_711.png""></a></li>"
        sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "promotion/s_conpri_riyou.asp""><img style=""width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src=""/img/top/clogo_familymart.png""></a></li>"
        sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "promotion/s_conpri_riyou.asp""><img style=""width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src=""/img/top/clogo_lawson.png""></a></li>"
		sHTML = sHTML & "</ul>"
		'</便利ツール>
	ElseIf iTabIndexType = 3 Then
		'<転職サポート>
		sHTML = sHTML & "<ul>"
		sHTML = sHTML & "<li class=""sidemenu_big"">転職サポート</li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "staff/jobcon/introduction.asp"">ジョブ・コンシェルジュ</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/jobcon/careeranalyzer/"">自己分析ツール</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "staff/jobcon/searchadvice/"">検索条件補助ツール</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/jobcon/interviewsimulator/"">面接対策</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "staff/notification_mail_service.asp"">スケジュール通知サービス</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_ready.asp"">転職の心構え</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_proce.asp"">転職に必要な手続き</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_goukaku.asp"">面接対策マニュアル</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_kyuuyomeisai.asp"">給与明細について</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/navistep_index.asp"">初めての転職活動</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "column/column_index.asp"">転職・就職コラム</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_mynavi.asp"">適職診断「じぶんナビ」</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/businesscolumns/"">ビジネスコラム</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_introduce.asp"">人材紹介</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_temporary.asp"">人材派遣</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_temptoperm.asp"">紹介予定派遣</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu_end""><a href=""" & HTTPS_CURRENTURL & "staff/jobcon/careerconsultation/"">キャリア相談</a></li>"
		sHTML = sHTML & "</ul>"
		'</転職サポート>
	ElseIf iTabIndexType = 4 Then
		'<コミュニティ>
		sHTML = sHTML & "<ul>"
		sHTML = sHTML & "<li class=""sidemenu_big"">コミュニティ</li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "cafe/cafe_list.asp"">しごとナビカフェ</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_introduce_swf.asp"">人材紹介劇</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "s_contents/s_mynavi.asp"">適職診断「じぶんナビ」</a></li>"
		sHTML = sHTML & "<li class=""sidemenu_end""><a href=""" & HTTP_CURRENTURL & "s_contents/enquete.asp"">しごとナビアンケート</a></li>"
		sHTML = sHTML & "</ul>"
		'</コミュニティ>
	ElseIf iTabIndexType = 5 Then
		'<採用ご担当者>
		sHTML = sHTML & "<ul>"
		sHTML = sHTML & "<li class=""sidemenu_big"">採用ご担当者</li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/about.asp"">しごとナビの特色</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_function.asp"">サービス概要</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "company/costperformance/"">採用改善サポート</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "company/request01.asp"">求人広告掲載申込み</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "staff/kiyaku.asp"">ご利用規約</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTPS_CURRENTURL & "company/access.asp"">お問合せ</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_introduce.asp"">人材紹介</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_dispatch.asp"">人材派遣</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_temptoperm.asp"">紹介予定派遣</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_successpoint.asp"">ご利用について</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_scout3point.asp"">スカウトメールのポイント</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/qa.asp"">採用Q&A</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "jinzaisearch/index.asp"">情報お試し検索</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/c_staffdata.asp"">求職者集計データ</a></li>"
		sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/research.asp"">採用方法診断</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu""><a href=""" & HTTP_CURRENTURL & "company/charge.asp"">ご利用プラン</a></li>"
		'sHTML = sHTML & "<li class=""sidemenu_end""><a href=""" & HTTP_CURRENTURL & "company/c_voice.asp"">ご利用企業さまの声</a></li>"
		sHTML = sHTML & "</ul>"
		'</採用ご担当者>
	ElseIf iTabIndexType = 6 Then
		'<My Page(求職者)>
		'</My Page(求職者)>
	ElseIf iTabIndexType = 7 Then
		'<My Page(企業)>
		'</My Page(企業)>
	ElseIf iTabIndexType = 8 Then	
	
		sHTML = sHTML &""
		

		
	End If


	Response.Write "</div>"'メインコンテンツの幅指定divの閉め（開始はheader最下部）

	If SidemenuType <> 9 Then

		Response.Write "<nav id=""side"">"

		If iTabIndexType = 0 Then
			Response.Write sScript
			Response.Write sHTML10
			Response.Write sHTML12
			Response.Write sHTML91
			Response.Write sHTML80 'リス社員募集
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
		<li class="sidemenu_big">転職サポート</li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_resume.asp">履歴書の自動作成/フォーマットのダウンロード</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_resume_kakikata.asp">履歴書の書き方</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_resume_qa.asp">履歴書Q＆A</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_careersheet.asp">職務経歴書の自動作成/フォーマットのダウンロード</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_careersheet_kakikata_1.asp">職務経歴書の書き方</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_taishokunegai.asp">退職願の書き方</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_year_calculation.asp">西暦・和暦/学歴計算</a></li>
        
        <li class="sidetitle">履歴書・職務経歴書の印刷</li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>promotion/conpripromotion.asp"><img style="width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src="/img/top/clogo_711.png"></a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>promotion/s_conpri_riyou.asp"><img style="width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src="/img/top/clogo_familymart.png"></a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>promotion/s_conpri_riyou.asp"><img style="width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src="/img/top/clogo_lawson.png"></a></li>
        
        <li class="sidetitle">転職サポート案内</li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_introduce.asp">人材紹介</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_temporary.asp">人材派遣</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_temptoperm.asp">紹介予定派遣</a></li>
		<!--<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/jobcon/careerconsultation/">キャリア相談</a></li>-->
        
        <li class="sidetitle">マップ</li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>type_map.asp">職種・業種別マップ</a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>area_map.asp">地域別マップ</a></li>	
        <li class="sidemenu_end"><a href="<%= HTTP_CURRENTURL %>keyword_map.asp">キーワードマップ</a></li>	
		</ul>
		<!-- <a href="/point/pr/" target="_blank"><img src="/img/neo/point_present.png"></a> -->

        <!--<div id="side_pickup">
			ここにPickUp
        </div>-->
			<!--<a href="http://www.shigotonavi.co.jp/order/order_detail.asp?OrderCode=J0066098" target="_self"><img src="/img/banner/jisya/tokyo_20120822.gif"></a>-->
			<%	
		ElseIf iTabIndexType = 2 Then
			Response.Write sScript
			Response.Write sHTML10
			Response.Write sHTML12
			Response.Write sHTML91
			Response.Write sHTML80 'リス社員募集
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
			Response.Write sHTML80 'リス社員募集
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
			Response.Write sHTML80 'リス社員募集
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
			Response.Write sHTML80 'リス社員募集
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
			
		ElseIf iTabIndexType = 8 Then '学ぶ
			Response.Write sHTML10
			Response.Write sHTML12
			%>
			
        <ul>
		<li class="sidemenu_big">自己分析</li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/jobcon/introduction.asp">ジョブ・コンシェルジュ</a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/jobcon/careeranalyzer/">自己分析ツール</a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/jobcon/searchadvice/">検索条件補助ツール</a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_kyuuyomeisai.asp">給与明細について</a></li>
		<li class="sidemenu_end"><a href="<%= HTTP_CURRENTURL %>staff/notification_mail_service.asp">スケジュール通知</a></li>
		</ul>
        
        <ul>
		<li class="sidemenu_big">スキルアップ</li>
		<li class="sidemenu_end"><a href="<%= HTTP_CURRENTURL %>staff/jobcon/interviewsimulator/">面接対策</a></li>
		</ul>
			
        <ul>
		<li class="sidemenu_big">転職ノウハウ</li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_ready.asp">転職の心構え</a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_proce.asp">転職に必要な手続き</a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_goukaku.asp">面接対策マニュアル</a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/navistep_index.asp">初めての転職活動</a></li>
		<li class="sidemenu_end"><a href="<%= HTTP_CURRENTURL %>column/column_index.asp">転職・就職コラム</a></li>
		</ul>	

       <!-- <ul>
		<li class="sidemenu_big">地方自治体ページ</li>
		<li class="sidemenu_end"><a href="<%= HTTP_CURRENTURL %>s_contents/s_localgoverment.asp">地方自治体ページ</a></li>
		</ul>	-->

        <ul>
		<li class="sidemenu_big">ビジネスコラム</li>
		<li class="sidemenu_end"><a href="<%= HTTP_CURRENTURL %>s_contents/businesscolumns/">ビジネスコラム</a></li>
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
	'<求人数、企業数、求職者数>

	Response.Write "<img src=""/img/top/countericon_order.gif"" alt=""求人数"" border=""0"" style=""margin:0px 2px;"">求人<span class=""cnt"">" & iOrderCnt & "</span>件&nbsp;"
	Response.Write "<img src=""/img/top/countericon_company.gif"" alt=""企業数"" border=""0"" style=""margin:0px 2px;"">企業<span class=""cnt"">" & iCompanyCnt & "</span>社&nbsp;"
	Response.Write "<img src=""/img/top/countericon_staff.gif"" alt=""求職者数"" border=""0"" style=""margin:0px 2px;"">求職者<span class=""cnt"">" & iAll & "</span>人&nbsp;"
	Response.Write "" & MonthName(Month(Now)) & Day(Now) & "日(" & Left(WeekdayName(Weekday(Now)),1) & ")" & "更新"
	'</求人数、企業数、求職者数>

%>
</p>
<a href="<%= HTTPS_CURRENTURL %>pr/pushpoint.asp"><img src="/img/common/how_to_use2.png"></a>
<a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp"><img src="/img/common/reg1_button_big_3.png" border="0" alt="会員登録(無料)" style="margin:0 0 0 15px;"></a>
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
<div style="width:120px; float:left;font-size: 15px; font-weight: bold; color: chocolate; text-align:center; line-height: 90px; border-radius:8px 0 0 8px;border-right: 2px dashed #ffd2a9;">ログイン</div>

<div class="right" style="float: left!important;margin: 11px 0 5px 25px;">

	<div class="left center" style="font-size: 14px;">
		ID
	<% If si <> "" Then %>
			<input type="text" name="CONF_UserID" value="<%= si %>" style="margin: 0 5px;width:120px; height: 25px; border-radius: 4px;">
	<%	Else %>
			<input type="text" name="CONF_UserID" value="<%= Request.Cookies("id_memory") %>" style="margin: 0 5px;width:120px;height: 25px;border-radius: 4px;">
	<%	End If %>
		パスワード<input type="password" name="CONF_Password" value="" style="margin: 0 0 0 5px;width:120px;height: 25px;border-radius: 4px;">
	</div>
<br clear="all">
	<div align="right" style="margin:0 0 5px 0;">
		<label><input type="checkbox" name="frmautologinflag" value="1">自動ﾛｸﾞｲﾝ</label>[<span style="color:#0045f9; cursor:pointer;" onclick="window.open('/infomation/autologin.asp','autologin','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=400,height=220');"><u>？</u></span>]
        
		<a href="<%= HTTPS_CURRENTURL %>staff/qa.asp#003" style="font-size:10px;">ﾛｸﾞｲﾝできない</a>
		<a href="<%= HTTPS_CURRENTURL %>staff/passwordreminder.asp" style="font-size:10px;">ID・ﾊﾟｽﾜｰﾄﾞを忘れた</a> 
		<input type="submit" value="ログイン" onclick="DataCheckIdreg(); return false" style="background: #FFA500;
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

		ElseIf iTabIndexType = 11 Then '交流
			Response.Write sHTML10
			Response.Write sHTML12	
			
			
		ElseIf iTabIndexType = 12 Then 'リンク
			Response.Write sHTML10
			Response.Write sHTML12
		
		%>	
		<ul>
		<li class="sidemenu_big">コンテンツ</li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_introduce_swf.asp">人材紹介劇</a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_mynavi.asp">適職診断「じぶんナビ」</a></li>
        <li class="sidemenu_end"><a href="<%= HTTP_CURRENTURL %>s_contents/enquete.asp">しごとナビアンケート</a></li>

		</ul>
		<%
        		
		ElseIf iTabIndexType = 13 Then '探す
			
			Response.Write sHTML12		
			
				%>
			
       <ul>
		<li class="sidemenu_big">転職サポート</li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_resume.asp">履歴書の自動作成/フォーマットのダウンロード</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_resume_kakikata.asp">履歴書の書き方</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_resume_qa.asp">履歴書Q＆A</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_careersheet.asp">職務経歴書の自動作成/フォーマットのダウンロード</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_careersheet_kakikata_1.asp">職務経歴書の書き方</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_taishokunegai.asp">退職願の書き方</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_year_calculation.asp">西暦・和暦/学歴計算</a></li>
        <li class="sidetitle">履歴書・職務経歴書の印刷</li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>promotion/conpripromotion.asp"><img style="width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src="/img/top/clogo_711.png"></a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>promotion/s_conpri_riyou.asp"><img style="width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src="/img/top/clogo_familymart.png"></a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>promotion/s_conpri_riyou.asp"><img style="width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src="/img/top/clogo_lawson.png"></a></li>
        <li class="sidetitle">転職サポート案内</li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_introduce.asp">人材紹介</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_temporary.asp">人材派遣</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_temptoperm.asp">紹介予定派遣</a></li>
		<!--<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/jobcon/careerconsultation/">キャリア相談</a></li>-->
	
		</ul>
        
        <%
ElseIf iTabIndexType = 14 Then
			Response.Write sHTML10
			Response.Write sHTML12		
			
				%>
			
       <ul>
		<li class="sidemenu_big">転職サポート</li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_resume.asp">履歴書の自動作成/フォーマットのダウンロード</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_resume_kakikata.asp">履歴書の書き方</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_resume_qa.asp">履歴書Q＆A</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_careersheet.asp">職務経歴書の自動作成/フォーマットのダウンロード</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/s_careersheet_kakikata_1.asp">職務経歴書の書き方</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_taishokunegai.asp">退職願の書き方</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_year_calculation.asp">西暦・和暦/学歴計算</a></li>
        <li class="sidetitle">履歴書・職務経歴書の印刷</li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>promotion/conpripromotion.asp"><img style="width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src="/img/top/clogo_711.png"></a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>promotion/s_conpri_riyou.asp"><img style="width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src="/img/top/clogo_familymart.png"></a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>promotion/s_conpri_riyou.asp"><img style="width:70px;border:0px solid #000;vertical-align:middle;margin-left:8px;"" src="/img/top/clogo_lawson.png"></a></li>
        <li class="sidetitle">転職サポート案内</li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_introduce.asp">人材紹介</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_temporary.asp">人材派遣</a></li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>s_contents/s_temptoperm.asp">紹介予定派遣</a></li>
		<!--<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>staff/jobcon/careerconsultation/">キャリア相談</a></li>-->
        <li class="sidetitle">マップ</li>
		<li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>type_map.asp">職種・業種別マップ</a></li>
        <li class="sidemenu"><a href="<%= HTTP_CURRENTURL %>area_map.asp">地域別マップ</a></li>	
        <li class="sidemenu_end"><a href="<%= HTTP_CURRENTURL %>keyword_map.asp">キーワードマップ</a></li>	
		</ul>

			<%	
			
				
		End If

		Response.Write "</nav>"
	End If
End Function
%>

