<%
'**********************************************************************************************************************
'概　要：企業情報向けの関数群
'　　　：
'　　　：■■■　前提条件　■■■
'　　　：要事前インクルード
'　　　：/config/personel.asp
'　　　：/include/commonfunc.asp
'一　覧：■■■　値取得用　■■■
'　　　：GetNaviContact			：対象企業が、対象求職者と過去にメールのやり取りがあるかどうかをチェック。
'　　　：ChkNaviRcvMail			：対象企業が、対象求職者から過去にメールを受け取ったことがあるかどうかをチェック。※ライセンス切れ後にメール送信可能かどうかを判定
'　　　：ChkMailAble			：メール送信可否を取得
'　　　：■■■　出力用　■■■
'　　　：DspScoutLimit			：企業のライセンス状況テーブルの表示, スカウト可否取得
'　　　：GetHtmlMyMenuAdvManager：企業マイメニューの求人広告管理部分ＨＴＭＬ取得
'　　　：GetHtmlMyMenuMyOrders	：企業マイメニューの自社求人票一覧
'　　　：■■■　チェック用　■■■
'　　　：ChkEditOrder			：求人票登録権限可否チェック
'**********************************************************************************************************************

'******************************************************************************
'概　要：対象企業が、対象求職者から過去にメールを受け取ったことがあるかどうかをチェック。
'引　数：rDB			：
'　　　：vCompanyCode	：企業コード
'　　　：vStaffCode		：求職者コード
'戻り値：Boolean		：[True]スカウト可能 [False]スカウト不可
'備　考：ライセンス切れ後にメール送信可能かどうかを判定
'使用元：しごとナビ/company/mailtoperson.asp
'更　新：2008/06/06 LIS K.kokubo 作成
'******************************************************************************
Function GetNaviContact(ByRef rDB, ByVal vCompanyCode, ByVal vStaffCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	GetNaviContact = False

	sSQL = "up_ChkNaviContact '" & vCompanyCode & "', '" & vStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		If oRS.Collect("ContactFlag") = "1" Then GetNaviContact = True
	End If
	Call RSClose(oRS)
End Function

'******************************************************************************
'概　要：対象企業が、対象求職者から過去にメールを受け取ったことがあるかどうかをチェック。
'引　数：rDB			：
'　　　：vCompanyCode	：企業コード
'　　　：vStaffCode		：求職者コード
'戻り値：Boolean		：[True]スカウト可能 [False]スカウト不可
'作成者：Lis Kokubo
'作成日：2007/02/18
'備　考：
'使用元：しごとナビ/company/mailtoperson.asp
'******************************************************************************
Function ChkNaviRcvMail(ByRef rDB, ByVal vCompanyCode, ByVal vStaffCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	ChkNaviRcvMail = False

	sSQL = "EXEC up_ChkNaviRcvMail '" & vCompanyCode & "', '" & vStaffCode & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		If oRS.Collect("RcvFlag") = "1" Then ChkNaviRcvMail = True
	End If
	Call RSClose(oRS)
End Function

'******************************************************************************
'概　要：対象企業が、対象求職者へメール送信可能か否かをチェック。
'引　数：rDB			：
'　　　：vCompanyCode	：企業コード
'　　　：vStaffCode		：求職者コード
'　　　：rMailAbleFlag	：[OUTPUT]メール送信可否フラグ [True]可 [False]不可
'　　　：rScoutFlag		：[OUTPUT]スカウトフラグ [True]スカウト [False]スカウトでない
'戻り値：String			：メール送信注意文言
'使用元：しごとナビ/company/mailtoperson.asp
'備　考：
'履　歴：2009/03/23 LIS K.Kokubo 作成
'******************************************************************************
Function ChkMailAble(ByRef rDB, ByVal vCompanyCode, ByVal vStaffCode, ByRef rMailAbleFlag, ByRef rScoutFlag)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbMailAbleFlag
	Dim dbScoutFlag
	Dim dbReceiveFlag
	Dim dbLicenseStatus
	Dim dbScoutLimitOverFlag
	Dim dbPlanTypeName

	rMailAbleFlag = False
	rScoutFlag = True
	If vCompanyCode = "" Or vStaffCode = "" Then Exit Function

	sSQL = ""
	sSQL = sSQL & "/* しごとナビ メール送信可否,スカウトフラグ取得 */" & vbCrLf
	sSQL = sSQL & "EXEC up_ChkMailAble '" & vCompanyCode & "', '" & vStaffCode & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		dbMailAbleFlag = oRS.Collect("MailAbleFlag")
		dbScoutFlag = oRS.Collect("ScoutFlag")
		dbReceiveFlag = oRS.Collect("ReceiveFlag")
		dbLicenseStatus = oRS.Collect("LicenseStatus")
		dbScoutLimitOverFlag = oRS.Collect("ScoutLimitOverFlag")
		dbPlanTypeName = oRS.Collect("PlanTypeName")

		If dbMailAbleFlag = "1" Then rMailAbleFlag = True
		If dbScoutFlag = "0" Then rScoutFlag = False

		If dbLicenseStatus = "public" Then
			'ライセンスが有効(発行日<=本日<=掲載終了日)
			If dbPlanTypeName = "mail" Then
				'メール課金プラン
			Else
				'メール課金プラン以外の場合は「スカウト」の考えを導入
				If dbScoutFlag = "1" Then
					'スカウト対象
					If dbScoutLimitOverFlag = "0" Then
						'スカウト制限ＯＫ
						ChkMailAble = "※スカウトの対象となる人物です。メールを送るとスカウトメールの送信数に数えます。"
					Else
						'スカウト制限ＮＯ
						ChkMailAble = "※スカウトの制限数を超えているため、この求職者へのメール送信はできません。"
					End If
				Else
					'非スカウト対象
					ChkMailAble = "※過去にメールのやりとりの実績が有る求職者です。メールを送ってもスカウトメールの送信数に数えません。"
				End If
			End If
		ElseIf dbLicenseStatus = "valid" Then
			'ライセンスが有効だけど、掲載日はまだ(発行日<=本日<=掲載開始日)
			ChkMailAble = "※掲載開始日に達していないため、まだメールを送信することはできません。"
		ElseIf dbLicenseStatus = "mailread" Then
			'ライセンスは切れているがメール閲覧期間中(掲載終了日<=本日<=掲載終了日+7日)
			If dbReceiveFlag = "0" Then
				'メールの受信実績が無い(スカウト対象とは別物である点に注意)
				ChkMailAble = "※メールの閲覧可能期間中の場合、メールの受信実績の無い求職者へメールを送信することはできません。"
			Else
				'メールの受信実績が有る(メール送信可能)
				ChkMailAble = "※過去にやりとりの実績が有る求職者です。ライセンス切れですがメール閲覧期間中であればメール可能です。"
			End If
		Else
			'ライセンス切れ
			ChkMailAble = "※ライセンスが切れているため、メールを送信することはできません。"
		End If
	End If
	Call RSClose(oRS)
End Function

'******************************************************************************
'概　要：企業のライセンス状況テーブルの表示, スカウト可否取得
'引　数：rDB			：
'　　　：vCompanyCode	：
'　　　：vDspOrderFlag	：
'戻り値：Boolean		：[True]スカウト可能 [False]スカウト不可
'備　考：
'使用元：しごとナビ/company/c_login.asp
'　　　：しごとナビ/dispatch/d_login.asp
'更　新：2007/02/13 LIS K.Kokubo 作成
'******************************************************************************
Function DspScoutLimit(ByRef rDB, ByVal vCompanyCode, ByVal vDspOrderFlag)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbScoutCnt
	Dim dbScoutLimit
	Dim dbMailUnitPrice
	Dim dbShimeFrom
	Dim dbShimeTo
	Dim dbHakouDate
	Dim dbRiyoFromDate
	Dim dbRiyoToDate
	Dim dbPermitOrderCnt
	Dim dbPublicOrderCnt
	Dim dbNotPublicOrderCnt
	Dim dbHalfwayOrderCnt
	Dim dbDspRiyoToDate
	Dim dbPointRemainder
	Dim dbPointWaiting
	Dim dbUsePoint
	Dim dbMailSendPayFlag
	Dim dbMailReceivePayFlag
	Dim dbMchFlag
	Dim dbSpMchFlag
	Dim dbPaySendMailPrice
	Dim dbPayReceiveMailPrice
	Dim dbPayMchPrice
	Dim dbPaySpMchNoticePrice
	Dim dbPaySpMchResponsePrice
	Dim dbPaySendMailCnt
	Dim dbPayReceiveMailCnt
	Dim dbPayMchCnt
	Dim dbPaySpMchNoticeCnt
	Dim dbPaySpMchResponseCnt

	Dim sHTML
	Dim sHTMLPrice
	Dim iScoutAble
	Dim sShimeFrom
	Dim sShimeTo
	Dim sHakouDate
	Dim sRiyoFromDate
	Dim sRiyoToDate
	Dim sDspRiyoToDate
	Dim iPrice

	sHTML = ""
	DspScoutLimit = False

	sSQL = sSQL & "/* しごとナビ 企業の利用状況取得 */" & vbCrLf
	sSQL = sSQL & "EXEC up_DtlUseStatusCompany_Advertisement '" & vCompanyCode & "';"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		dbScoutCnt = ChkStr(oRS.Collect("ScoutCnt"))
		dbScoutLimit = ChkStr(oRS.Collect("ScoutLimit"))
		dbMailUnitPrice = oRS.Collect("MailUnitPrice")
		dbShimeFrom = oRS.Collect("ShimeFrom")
		dbShimeTo = oRS.Collect("ShimeTo")
		dbHakouDate = ChkStr(oRS.Collect("HakouDate"))
		dbRiyoFromDate = ChkStr(oRS.Collect("RiyoFromDate"))
		dbRiyoToDate = ChkStr(oRS.Collect("RiyoToDate"))
		dbPermitOrderCnt = ChkStr(oRS.Collect("PermitOrderCnt"))
		dbPublicOrderCnt = ChkStr(oRS.Collect("PublicOrderCnt"))
		dbNotPublicOrderCnt = ChkStr(oRS.Collect("NotPublicOrderCnt"))
		dbHalfwayOrderCnt = ChkStr(oRS.Collect("HalfwayOrderCnt"))
		dbDspRiyoToDate = ChkStr(oRS.Collect("DspRiyoToDate"))
		dbPointRemainder = ChkStr(oRS.Collect("PointRemainder"))
		dbPointWaiting = oRS.Collect("PointWaiting")
		dbUsePoint = oRS.Collect("UsePoint")
		dbMailSendPayFlag = oRS.Collect("MailSendPayFlag")
		dbMailReceivePayFlag = oRS.Collect("MailReceivePayFlag")
		dbMchFlag = oRS.Collect("MchFlag")
		dbSpMchFlag = oRS.Collect("SpMchFlag")
		dbPaySendMailPrice = oRS.Collect("PaySendMailPrice")
		dbPayReceiveMailPrice = oRS.Collect("PayReceiveMailPrice")
		dbPayMchPrice = oRS.Collect("PayMchPrice")
		dbPaySpMchNoticePrice = oRS.Collect("PaySpMchNoticePrice")
		dbPaySpMchResponsePrice = oRS.Collect("PaySpMchResponsePrice")
		dbPaySendMailCnt = oRS.Collect("PaySendMailCnt")
		dbPayReceiveMailCnt = oRS.Collect("PayReceiveMailCnt")
		dbPayMchCnt = oRS.Collect("PayMchCnt")
		dbPaySpMchNoticeCnt = oRS.Collect("PaySpMchNoticeCnt")
		dbPaySpMchResponseCnt = oRS.Collect("PaySpMchResponseCnt")
	End If
	Call RSClose(oRS)

	iScoutAble = dbScoutLimit - dbScoutCnt
	sShimeFrom = Year(dbShimeFrom) & "年" & Month(dbShimeFrom) & "月" & Day(dbShimeFrom) & "日"
	sShimeTo = Year(dbShimeTo) & "年" & Month(dbShimeTo) & "月" & Day(dbShimeTo) & "日"
	sHakouDate = Year(dbHakouDate) & "年" & Month(dbHakouDate) & "月" & Day(dbHakouDate) & "日"
	sRiyoFromDate = Year(dbRiyoFromDate) & "年" & Month(dbRiyoFromDate) & "月" & Day(dbRiyoFromDate) & "日"
	sRiyoToDate = Year(dbRiyoToDate) & "年" & Month(dbRiyoToDate) & "月" & Day(dbRiyoToDate) & "日"
	sDspRiyoToDate = Year(dbDspRiyoToDate) & "年" & Month(dbDspRiyoToDate) & "月" & Day(dbDspRiyoToDate) & "日"

	DspScoutLimit = True

	sHTML = sHTML & "<div style=""margin-bottom:15px;"">"
	sHTML = sHTML & "<table style=""margin:0px;"">"
	sHTML = sHTML & "<colgroup>"
	sHTML = sHTML & "<col style=""width:112px;padding:3px;background-color:#e8e8ff;""></col>"
	sHTML = sHTML & "<col style=""width:473px;padding:3px;""></col>"
	sHTML = sHTML & "</colgroup>"
	sHTML = sHTML & "<tbody>"

	iPrice = (dbPaySendMailPrice + dbPayReceiveMailPrice + dbPayMchPrice + dbPaySpMchNoticePrice + dbPaySpMchResponsePrice) - dbUsePoint * 100

	'<料金>
	If InStr(dbMailSendPayFlag & dbMailReceivePayFlag & dbMchFlag & dbSpMchFlag, "1") > 0 Then
		sHTMLPrice = ""
		sHTMLPrice = sHTMLPrice & "<div style=""float:left;width:39%;"">"
		sHTMLPrice = sHTMLPrice & "現在の料金計&nbsp;:&nbsp;<b><span style=""color:#ff0000;"">" & GetJapaneseYen(iPrice) & "</span></b>"
		If dbUsePoint > 0 Then sHTMLPrice = sHTMLPrice & "&nbsp;(" & dbUsePoint & "pt利用：" & dbUsePoint * 100 & "円割引)"
		sHTMLPrice = sHTMLPrice & "</div>"

		sHTMLPrice = sHTMLPrice & "<div style=""float:right;width:59%;"">"
		If (dbMailSendPayFlag = "1" Or dbMailReceivePayFlag = "1") Or (dbPaySendMailCnt + dbPayReceiveMailCnt > 0) Then
			sHTMLPrice = sHTMLPrice & "課金メール送信数&nbsp;:&nbsp;<b><span style=""color:#0000ff;"">" & dbPaySendMailCnt & "名</span></b>(<b><span style=""color:#ff0000;"">" & GetJapaneseYen(dbPaySendMailPrice) & "</span></b>)<br>"
			sHTMLPrice = sHTMLPrice & "課金メール受信数&nbsp;:&nbsp;<b><span style=""color:#0000ff;"">" & dbPayReceiveMailCnt & "名</span></b>(<b><span style=""color:#ff0000;"">" & GetJapaneseYen(dbPayReceiveMailPrice) & "</span></b>)<br>"
		End If

		If dbMchFlag = "1" Or dbPayMchCnt > 0 Then
			sHTMLPrice = sHTMLPrice & "マッチング人材応募数&nbsp;:&nbsp;<b><span style=""color:#0000ff;"">" & dbPayMchCnt & "名</span></b>(<b><span style=""color:#ff0000;"">" & GetJapaneseYen(dbPayMchPrice) & "</span></b>)<br>"
		End If

		If dbSpMchFlag = "1" Or (dbPaySpMchNoticeCnt + dbPaySpMchResponseCnt > 0) Then
			sHTMLPrice = sHTMLPrice & "通知メール数&nbsp;:&nbsp;<b><span style=""color:#0000ff;"">" & dbPaySpMchNoticeCnt & "通</span></b>(<b><span style=""color:#ff0000;"">" & GetJapaneseYen(dbPaySpMchNoticePrice) & "</span></b>)<br>"
			sHTMLPrice = sHTMLPrice & "応募数&nbsp;:&nbsp;<b><span style=""color:#0000ff;"">" & dbPaySpMchResponseCnt & "名</span></b>(<b><span style=""color:#ff0000;"">" & GetJapaneseYen(dbPaySpMchResponsePrice) & "</span></b>)<br>"
		End If

		sHTMLPrice = sHTMLPrice & "</div>"
		sHTMLPrice = sHTMLPrice & "<div style=""clear:both;""></div>"
		sHTMLPrice = sHTMLPrice & "<div class=""line1""></div>"
	End If
	'</料金>

	If G_PLANTYPE = "mail" Then

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th style=""border:1px solid #cccccc;"">課金状況</th>"
		sHTML = sHTML & "<td style=""border:1px solid #cccccc;"">"

		sHTML = sHTML & sHTMLPrice

		sHTML = sHTML & "<span style=""font-size:10px;"">※算出期間：" & sShimeFrom & "&nbsp;～&nbsp;" & sShimeTo & "〆</span>"
		If Date < "2009/08/01" Then sHTML = sHTML & "<br>※<span style=""color:#ff0000;"">８月１日以降の受信メールが課金されるようになります。</span>"

		If vDspOrderFlag = True Then
			'メール課金プランの場合は利用状況へのリンクを表示
			sHTML = sHTML & "<p class=""m0""><a href=""/company/license/mailplan_status.asp"">→課金状況の過去分を確認</a>&nbsp;...&nbsp;過去の料金の明細などを確認できます。</p>"
		End If

		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"
	Else
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th style=""border:1px solid #cccccc;"">スカウトメール</th>"
		sHTML = sHTML & "<td style=""border:1px solid #cccccc;"">"
		sHTML = sHTML & "スカウト可能数：<b><span style=""color:#0000ff;"">" & iScoutAble & "件</span></b>&nbsp;&nbsp;"
		sHTML = sHTML & "スカウト送信数：" & dbScoutCnt & "件／最大" & dbScoutLimit & "件まで"

		If sHTMLPrice <> "" Then
			sHTML = sHTML & "<div class=""line1""></div>"
			sHTML = sHTML & sHTMLPrice
		End If

		If dbMchFlag = "1" Or dbSpMchFlag = "1" Then
			sHTML = sHTML & "<p class=""m0""><a href=""/company/license/mailplan_status.asp"">→課金状況の過去分を確認</a>&nbsp;...&nbsp;月額利用料以外の過去の料金の明細などを確認できます。</p>"
		End If

		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"
	End If

	If vDspOrderFlag = True Then
		If G_PLANTYPE = "mail" Then
			sHTML = sHTML & "<tr>"
			sHTML = sHTML & "<th style=""border:1px solid #cccccc;"">ポイント状況</th>"
			sHTML = sHTML & "<td style=""border:1px solid #cccccc;"">"
			sHTML = sHTML & "総ポイント&nbsp;:&nbsp;<b><span style=""color:Red;"">" & dbPointWaiting + dbPointRemainder & "pt</span></b>&nbsp;&nbsp;"
			sHTML = sHTML & "（利用可能ポイント&nbsp;:&nbsp;<b><span style=""color:#0000ff;"">" & dbPointRemainder & "pt</span></b>）<br>"
			sHTML = sHTML & "<span style=""font-size:10px;"">※ポイントは発生日から２ヵ月後に利用可能ポイントとなります。</span>"
			sHTML = sHTML & "<p class=""m0""><a href=""/company/point/"" style="""">→ポイント管理</a>&nbsp;...&nbsp;ポイントの残数の確認や利用申請・取消ができます。</p>"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "</tr>"
		End If

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th style=""border:1px solid #cccccc;"">掲載期間</th>"
		sHTML = sHTML & "<td style=""border:1px solid #cccccc;"">"
		sHTML = sHTML & sRiyoFromDate & "～" & sDspRiyoToDate
		If G_PLANTYPE = "mail" Then
			sHTML = sHTML & "<br>"
			sHTML = sHTML & "<span style=""font-size:10px;"">※ログインをすると、掲載終了日がその日より２ヵ月後の〆日(月末)に更新されます。</span>"
		End If
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th style=""border:1px solid #cccccc;"">求人票公開状況</th>"
		sHTML = sHTML & "<td style=""border:1px solid #cccccc;"">"
		sHTML = sHTML & "<p class=""m0"">"
		sHTML = sHTML & "掲載中(&nbsp;" & dbPublicOrderCnt & "件&nbsp;)&nbsp;&nbsp;"
		sHTML = sHTML & "非掲載(&nbsp;" & dbNotPublicOrderCnt & "件&nbsp;)&nbsp;&nbsp;"
		sHTML = sHTML & "審査中(&nbsp;" & dbPermitOrderCnt & "件&nbsp;)&nbsp;&nbsp;"
		sHTML = sHTML & "作成中(&nbsp;" & dbHalfwayOrderCnt & "件&nbsp;)&nbsp;&nbsp;"
		sHTML = sHTML & "</p>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"
	End If

	sHTML = sHTML & "</tbody>"
	sHTML = sHTML & "</table>"
	sHTML = sHTML & "</div>"

	Response.Write sHTML
End Function

'******************************************************************************
'概　要：企業マイメニューの求人広告管理部分ＨＴＭＬ取得
'引　数：rDB			：接続中ＤＢオブジェクト
'　　　：vCompanyCode	：企業コード
'　　　：vCompanyKbn	：企業種類
'　　　：vLicenseFlag	：ライセンス状況フラグ ["1"]利用中
'　　　：vCollectionCnt	：掲載中求人票件数
'戻り値：String
'使用元：しごとナビ/company/c_login.asp
'　　　：しごとナビ/dispatch/d_login.asp
'備　考：
'更　新：2007/12/04 LIS K.Kokubo 作成
'　　　：2008/01/31 LIS K.Kokubo プラチナプラン、ゴールドプランで外部からのアクセスの場合では「求人票の新規作成」リンクを非表示
'　　　：2009/03/11 LIS K.Kokubo 「求人票の新規作成」リンクを全てのプランに開放
'　　　：2009/06/25 LIS K.Kokubo 求人票検索条件保存追加
'　　　：2009/07/02 LIS K.Kokubo 一括メール管理追加
'******************************************************************************
Function GetHtmlMyMenuAdvManager(ByRef rDB, ByVal vCompanyCode, ByVal vCompanyKbn, ByVal vLicenseFlag, ByVal vPlanType, ByVal vCollectionCnt)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim dbCnt
	Dim sHTML
	Dim flgNew

	sHTML = ""

	If vLicenseFlag = "1" Then
		'ナビ求人広告ライセンスを利用中
		sHTML = sHTML & "<table border=""0"" style=""width:100%;"">"

		'sHTML = sHTML & "<tr>"
		'sHTML = sHTML & "<td colspan=""2"" bgcolor=""#666699""><font color=""#FFFFFF"">求人広告</font></td>"
		'sHTML = sHTML & "</tr>"

		'<「求人票の新規作成」リンク表示>
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td>"

		flgNew = True
		sSQL = "SELECT COUNT(*) AS Cnt FROM C_Info AS A INNER JOIN C_SupplementInfo AS B ON A.OrderCode = B.OrderCode WHERE RegistCommit = '0' AND A.CompanyCode = '" & G_USERID & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			dbCnt = oRS.Collect("Cnt")
			If dbCnt >= 5 Then flgNew = False
		End If
		If flgNew = True Then
			sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/orderedit/new.asp"">"
			sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"" border=""0"" alt="""">"
			sHTML = sHTML & "</a>"
		Else
			sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"" border=""0"" alt="""">"
		End If

		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td bgcolor=""#E8E8FF"" style=""border-bottom:1px solid #ffffff;"">"

		If flgNew = True Then
			'作成途中求人票数が規定以内なら新規作成のリンクを表示
			sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/orderedit/new.asp"">求人票の新規作成</a>"
		Else
			'作成途中求人票数が規定オーバーなら新規作成のリンクを表示
			sHTML = sHTML & "<p class=""m0"">求人票の新規作成</p>"
		End If

		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td style=""border-bottom:1px solid #ffffff;"">"
		If dbCnt >= 5 Then
			sHTML = sHTML & "<p class=""m0"" style=""line-height:16px;"">※作成途中の求人票は<b>５個</b>まで持つ事ができます。新たに求人票を作成したい場合は、作成途中の求人票一覧の「編集」ボタンより求人票を確定するか、「削除」ボタンで削除する必要があります。</p>"
		Else
			sHTML = sHTML & "新しい求人票を作成します。【<a href=""http://jinzai.shigotonavi.co.jp/joboffer/make_advertisement.asp"" target=""blank_"">求人票作成のポイント</a>】<br>"
		End If
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"
		'</「求人票の新規作成」リンク表示>

		If vCollectionCnt = 0 Then
			'募集中の求人が無い場合
			sHTML = sHTML & "<tr>"
			sHTML = sHTML & "<td>"
			sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"">"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "<td bgcolor=""#E8E8FF"" style=""border-bottom:1px solid #ffffff;"">"
			sHTML = sHTML & "－－－"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "<td>"
			sHTML = sHTML & "現在、求人募集している求人票が無いため、求職者検索はご利用できません。<br>"
			sHTML = sHTML & "あらたに求人票を作成いただければ、弊社にて確認の上、ご利用いただけます。"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "</tr>"
		Else
			'募集中の求人がある場合

			'求人票のコピー作成
			sHTML = sHTML & "<tr>"
			sHTML = sHTML & "<td>"
			sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/myorderlist.asp"">"
			sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"" border=""0"" alt="""">"
			sHTML = sHTML & "</a>"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "<td>"
			sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/myorderlist.asp"">求人票のコピー作成</a>"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "<td>"
			sHTML = sHTML & "御社の既存の求人票を元に新しい求人票を作成できます。"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "</tr>"

			'求人票修正
			sHTML = sHTML & "<tr>"
			sHTML = sHTML & "<td>"
			sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/myorderlist.asp"">"
			sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"" border=""0"" alt="""">"
			sHTML = sHTML & "</a>"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "<td>"
			sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/myorderlist.asp"">求人票修正</a>"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "<td>現在御社にて募集している求人票の検索と修正ができます。</td>"
			sHTML = sHTML & "</tr>"

			'求職者検索
			sHTML = sHTML & "<tr>"
			sHTML = sHTML & "<td>"
			sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/myorderlist.asp"">"
			sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"" border=""0"" alt="""">"
			sHTML = sHTML & "</a>"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "<td>"
			sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/myorderlist.asp"">求職者の検索とスカウト</a>"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "<td>求職者を検索し、スカウトできます。</td>"
			sHTML = sHTML & "</tr>"

			'求職者検索条件管理
			sHTML = sHTML & "<tr>"
			sHTML = sHTML & "<td>"
			sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/searchstaffcondition/list.asp"">"
			sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"" border=""0"" alt="""">"
			sHTML = sHTML & "</a>"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "<td>"
			sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/searchstaffcondition/list.asp"">求職者検索条件管理</a>"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "<td>保存した求職者検索条件を削除・名称変更します。</td>"
			sHTML = sHTML & "</tr>"
		End If

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td colspan=""3"">"
		sHTML = sHTML & "<div class=""line1""></div>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"


		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/company_reg1.asp"">"
		sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"" border=""0"" alt="""">"
		sHTML = sHTML & "</a>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/company_reg1.asp"">自社情報更新</a>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td>求人情報以外の部分、御社概要の編集ができます。</td>"
		sHTML = sHTML & "</tr>"

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTP_CURRENTURL & "company/img_upload.asp"">"
		sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"" border=""0"" alt="""">"
		sHTML = sHTML & "</a>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTP_CURRENTURL & "company/img_upload.asp"">企業写真掲載</a>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "企業紹介に利用する、代表的な画像（ロゴなど）を登録できます。"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTP_CURRENTURL & "company/company_img_list.asp"">"
		sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"" border=""0"" alt="""">"
		sHTML = sHTML & "</a>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTP_CURRENTURL & "company/company_img_list.asp"">求人票用画像ストック</a>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "求人票に複数画像を載せる場合は、ここで画像を登録しておきます。"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/mailhistory_company.asp"">"
		sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"" border=""0"" alt="""">"
		sHTML = sHTML & "</a>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "company/mailhistory_company.asp"">メール履歴</a>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td>求職者とのメール履歴や、採用の進捗管理が可能。</td>"
		sHTML = sHTML & "</tr>"

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTP_CURRENTURL & "company/lumpmail/list.asp"">"
		sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"" border=""0"" alt="""">"
		sHTML = sHTML & "</a>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTP_CURRENTURL & "company/lumpmail/list.asp"">一括メール管理</a>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td style=""border-bottom:1px solid #ffffff;"">一括メールの予約状況の確認、一括メールの作成・送信が可能。</td>"
		sHTML = sHTML & "</tr>"

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTP_CURRENTURL & "mailtemplate/manager.asp"">"
		sHTML = sHTML & "<img src=""/img/6.gif"" width=""20"" height=""13"" border=""0"" alt="""">"
		sHTML = sHTML & "</a>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTP_CURRENTURL & "mailtemplate/manager.asp"">メールテンプレート管理</a>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td>メールを作成する際に利用できる雛形を管理します。【<a href=""/company/c_scout3point.asp"">スカウトメール作成のポイント</a>】</td>"
		sHTML = sHTML & "</tr>"
		sHTML = sHTML & "</table>"
	End If
	'******************************************************************************
	'** 求人広告 end
	'******************************************************************************

	GetHtmlMyMenuAdvManager = sHTML
End Function

'******************************************************************************
'概　要：未確定の求人票一覧
'引　数：rDB			：接続中ＤＢ
'　　　：vCompanyCode	：企業コード
'戻り値：String			：自社求人票一覧ＨＴＭＬ
'備　考：
'使用元：しごとナビ/company/mailtoperson.asp
'更　新：2008/11/04 LIS K.Kokubo 作成
'******************************************************************************
Function GetHtmlUnCommitOrders(ByRef rDB, ByVal vCompanyCode)
	'<変数宣言>
	Dim sHTML

	Dim dbOrderCode
	Dim dbJobTypeDetail
	Dim dbUpdateDay
	'</変数宣言>

	'<変数初期化>
	sHTML = ""
	'</変数初期化>

	sSQL = "SELECT A.OrderCode, A.JobTypeDetail, A.UpdateDay FROM C_Info AS A INNER JOIN C_SupplementInfo AS B ON A.OrderCode = B.OrderCode WHERE RegistCommit = '0' AND A.CompanyCode = '" & G_USERID & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		sHTML = sHTML & "<table class=""pattern3"" style=""width:100%;"">"
		sHTML = sHTML & "<thead>"
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th style=""width:150px"">更新日時</th>"
		sHTML = sHTML & "<th style=""width:150px"">有効期限</th>"
		sHTML = sHTML & "<th>職種</th>"
		sHTML = sHTML &"</tr>"
		sHTML = sHTML &"</thead>"
		sHTML = sHTML & "<tbody>"

		Do While GetRSState(oRS) = True
			dbOrderCode = oRS.Collect("OrderCode")
			dbJobTypeDetail = oRS.Collect("JobTypeDetail")
			dbUpdateDay = oRS.Collect("UpdateDay")

			sHTML = sHTML & "<tr>"
			sHTML = sHTML & "<td>" & GetDateStr(dbUpdateDay, "/") & "<br>" & GetTimeStr(dbUpdateDay, ":") & "</td>"
			sHTML = sHTML & "<td>" & GetDateStr(DateAdd("d", 6, dbUpdateDay), "/") & "<br>00:00:00</td>"
			sHTML = sHTML & "<td>"
			sHTML = sHTML & "<form action=""/company/orderedit/base.asp?ordercode="& dbOrderCode & """ method=""post"" style=""display:inline;""><input class=""btn1"" type=""submit"" value=""編集""></form>&nbsp;"
			sHTML = sHTML & "<form action="""" method=""post"" style=""display:inline;"" onsubmit=""return confirm('削除しますか？');""><input name=""frmdelordercode"" type=""hidden"" value=""" & dbOrderCode & """><input class=""btn1"" type=""submit"" value=""削除""></form>&nbsp;"
			sHTML = sHTML & "<input type=""text"" value=""" & dbJobTypeDetail & """ style=""width:350px; border-width:0px; background-color:transparent;"">"
			sHTML = sHTML &"</td>"
			sHTML = sHTML &"</tr>"

			oRS.MoveNext
		Loop

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td style=""width:60px; padding:0px; border-width:0px;""></td>"
		sHTML = sHTML & "<td style=""width:60px; padding:0px; border-width:0px;""></td>"
		sHTML = sHTML & "<td style=""width:480px; padding:0px; border-width:0px;""></td>"
		sHTML = sHTML & "</tr>"
		sHTML = sHTML & "</tbody>"
		sHTML = sHTML & "</table>"
	End If

	GetHtmlUnCommitOrders = sHTML
End Function

'******************************************************************************
'概　要：企業マイメニューの自社求人票一覧
'引　数：rDB			：接続中ＤＢ
'　　　：vUserType		：ログインユーザ種類
'　　　：vCompanyCode	：企業コード
'　　　：vPageSize		：１ページあたりの最大出力件数
'　　　：vPage			：出力ページ
'　　　：vSort			：データ並び替え
'　　　：vPersonName	：絞込み求人担当者
'戻り値：String			：自社求人票一覧ＨＴＭＬ
'備　考：
'使用元：しごとナビ/company/c_login.asp
'履　歴：2007/02/18 LIS K.Kokubo 作成
'　　　：2009/05/22 LIS K.Kokubo 自社求人一覧テーブルのテーブルヘッダを非表示。（未読メール、未読求職者の説明が読まれていないかも対策）
'******************************************************************************
Function GetHtmlMyMenuMyOrders(ByRef rDB, ByVal vCompanyCode, ByVal vUserType, ByVal vPageSize, ByVal vPage, ByVal vSort, ByVal vPersonName)
	'<変数宣言>
	Dim sSQL
	Dim oRS
	Dim oRS2
	Dim oRS3
	Dim flgQE
	Dim sError

	Dim dbOrderCode		'情報コード
	Dim dbJobTypeDetail	'具体的職種名
	Dim dbPersonName	'求人担当者
	Dim dbMailCnt		'未読メール数
	Dim dbStaffCnt		'新着求職者数
	Dim dbPublicFlag	'掲載状態フラグ ["1"]掲載中 ["0"]非掲載
	Dim dbSecretFlag	'シークレット求人フラグ ["1"]シークレット求人 ["0"]一般公開求人
	Dim dbSearchName	'検索条件名
	Dim dbSearchParam	'検索条件パラメータ

	Dim sHTML
	Dim sHTML2
	Dim sPageControl
	Dim sURL
	Dim sParam
	Dim sJobTypeDetail
	Dim sMailCnt
	Dim sStaffCnt
	Dim iRow			'自社求人票の出力中レコード番号
	Dim sOnChange
	'保存検索条件取得用
	Dim tmpAbsolutePage
	Dim sXML
	'</変数宣言>

	'ＵＲＬ
	If vUserType = "company" Then
		sURL = "/company/c_login.asp"
	ElseIf vUserType = "dispatch" Then
		sURL = "/dispatch/d_login.asp"
	End If

	sParam = ""
	'並び替えパラメータ
	If vSort <> "" Then
		If sParam <> "" Then sParam = sParam & "&amp;"
		sParam = sParam & "sort=" & vSort
	End If
	If vPersonName <> "" Then
		If sParam <> "" Then sParam = sParam & "&amp;"
		sParam = sParam & "pn=" & Server.URLEncode(vPersonName)
	End If
	If sParam <> "" Then sParam = "?" & sParam

	'自社求人票一覧
	sHTML = ""
	sSQL = "up_LstMyMenuMyOrders '" & vCompanyCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	If GetRSState(oRS) = False Then Exit Function

	oRS.PageSize = vPageSize
	If vPersonName <> "" Then oRS.Filter = "PersonName = '" & vPersonName & "'"
	If GetRSState(oRS) = True Then
		If vPage <> "" Then oRS.AbsolutePage = vPage
		Select Case vSort
			Case Else: oRS.Sort = "PublicFlag DESC, OrderCode DESC"
		End Select
	End If

	'ページコントロール取得
	sPageControl = GetHtmlPageControlParam(rDB, oRS, vPageSize, vPage, sURL & sParam, "myorder")

	'<保存検索条件取得>
	tmpAbsolutePage = oRS.AbsolutePage
	iRow = 1
	sXML = "<root>"
	Do While GetRSState(oRS) = True And iRow <= vPageSize
		dbOrderCode = oRS.Collect("OrderCode")
		sXML = sXML & "<order><ordercode>" & dbOrderCode & "</ordercode></order>"
		oRS.MoveNext
	Loop
	sXML = sXML & "</root>"

	sSQL = ""
	sSQL = sSQL & "/* 保存検索条件取得 */" & vbCrLf
	sSQL = sSQL & "EXEC up_LstCMPSearchStaffCondition_XML '" & sXML & "';"
	flgQE = QUERYEXE(dbconn, oRS2, sSQL, sError)
	If GetRSState(oRS2) = True Then
		Set oRS2.ActiveConnection = Nothing
	End If
	oRS.AbsolutePage = tmpAbsolutePage
	'</保存検索条件取得>

	iRow = 1
	Do While GetRSState(oRS) = True And iRow <= vPageSize
		dbOrderCode = oRS.Collect("OrderCode")
		dbJobTypeDetail = oRS.Collect("JobTypeDetail")
		dbPersonName = oRS.Collect("PersonName")
		dbPublicFlag = oRS.Collect("PublicFlag")
		dbSecretFlag = oRS.Collect("SecretFlag")

		sJobTypeDetail = dbJobTypeDetail
		If Len(sJobTypeDetail) > 29 Then sJobTypeDetail = Left(sJobTypeDetail, 29) & "..."

		'未読メール数
		sSQL = "up_CntMyMenuNotReadMail '" & dbOrderCode & "'"
		flgQE = QUERYEXE(rDB, oRS3, sSQL, sError)
		If GetRSState(oRS3) = True Then
			dbMailCnt = oRS3.Collect("Cnt")
			If dbMailCnt > 0 Then
				sMailCnt = "<div class=""iconred"">未読メール</div>&nbsp;"
				sMailCnt = sMailCnt & "<a href=""" & HTTP_CURRENTURL & "company/mailhistory_company.asp?soc=" & dbOrderCode & """>" & dbMailCnt & "件</a>"
			Else
				sMailCnt = "<div class=""icongray"">未読メール</div>&nbsp;<span style=""color:#999999;"">なし</span>"
			End If
		End If
		Call RSClose(oRS3)

		'新着求職者数
		sSQL = "up_CntMyMenuNewStaff '" & dbOrderCode & "'"
		flgQE = QUERYEXE(rDB, oRS3, sSQL, sError)
		If GetRSState(oRS3) = True Then
			dbStaffCnt = oRS3.Collect("Cnt")
			If dbStaffCnt > 0 Then
				sStaffCnt = "<div class=""iconred"">未読求職者</div>&nbsp;"
				sStaffCnt = sStaffCnt & "<a href=""" & HTTP_CURRENTURL & "staff/person_list.asp?ordercode=" & dbOrderCode & "&amp;rdfrom=" & GetDateStr(DateAdd("d", -9, Date), "") & """>" & dbStaffCnt & "人</a>"
			Else
				sStaffCnt = "<div class=""icongray"">未読求職者</div>&nbsp;<span style=""color:#999999;"">なし</span>"
			End If
		End If
		Call RSClose(oRS3)

		sHTML = sHTML & "<tr>"
		'<求人>
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "order/order_detail.asp?ordercode=" & dbOrderCode & """>" & sJobTypeDetail & "</a>&nbsp;"
		sHTML = sHTML & "(担当：" & dbPersonName & ")<br>"
		If dbPublicFlag = "1" Then
			sHTML = sHTML & "<div class=""iconredbg"">掲載中</div>&nbsp;"
		ElseIf dbPublicFlag = "0" Then
			sHTML = sHTML & "<div class=""icongraybg"">非掲載</div>&nbsp;"
		End If
		If dbSecretFlag = "1" Then sHTML = sHTML & "<div class=""icongraybg"">SECRET</div>&nbsp;"
		sHTML = sHTML & sMailCnt & "&nbsp;"
		sHTML = sHTML & sStaffCnt

		sHTML = sHTML & "</td>"
		'</求人>

		'<求職者検索>
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<div style=""margin-bottom:3px;"">"
		sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""自動検索"" onclick=""location.href='/staff/person_list.asp?ordercode=" & dbOrderCode & "';"">&nbsp;"
		sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""詳細検索"" onclick=""location.href='/staff/person_search_detail.asp?ordercode=" & dbOrderCode & "&amp;setdata=1';"">"
		sHTML = sHTML & "</div>"

		'<保存検索条件>
		If GetRSState(oRS2) = True Then
			oRS2.Filter = 0
			oRS2.Filter = "OrderCode = '" & dbOrderCode & "'"
			If GetRSState(oRS2) = True Then
				sHTML = sHTML & "<div class=""line1""></div>"
				sHTML = sHTML & "<ul>"
				Do While GetRSState(oRS2) = True
					dbSearchName = oRS2.Collect("SearchName")
					dbSearchParam = oRS2.Collect("SearchParam")

					sHTML = sHTML & "<li>・<a href=""" & HTTP_CURRENTURL & "staff/person_list.asp?ordercode=" & dbOrderCode & "&amp;" & dbSearchParam & """>" & dbSearchName & "</a></li>"

					oRS2.MoveNext
				Loop
				sHTML = sHTML & "</ul>"
			End If
			'</保存検索条件>
		End If
		sHTML = sHTML & "</td>"
		'</求職者検索>
		sHTML = sHTML & "</tr>"

		iRow = iRow + 1
		oRS.MoveNext
	Loop
	Call RSClose(oRS2)
	Call RSClose(oRS)

	sHTML2 = ""
	If sHTML <> "" Then
		sOnChange = "location.href='" & sURL & "?sort=" & vSort & "&amp;pn=' + escape(this.value) + '#myorder';"
		sHTML2 = sHTML2 & "<div id=""myorder""></div>"
		sHTML2 = sHTML2 & sPageControl & vbCrLf
		sHTML2 = sHTML2 & "<table class=""pattern3"" border=""0"" style=""width:100%;"">"
		'sHTML2 = sHTML2 & "<thead>"
		'sHTML2 = sHTML2 & "<tr>"
		'sHTML2 = sHTML2 & "<th colspan=""3"">自社求人票一覧</th>"
		'sHTML2 = sHTML2 & "</tr>"
		'sHTML2 = sHTML2 & "</thead>"
		sHTML2 = sHTML2 & "<thead>"
		sHTML2 = sHTML2 & "<tr>"
		sHTML2 = sHTML2 & "<th style=""width:388px;"">職種&nbsp;<select name=""pn"" onchange=""" & sOnChange & """><option value="""">--求人担当--</option>" & GetContactPersonNameOptionHtml(vCompanyCode, vPersonName) & "</select></th>"
		sHTML2 = sHTML2 & "<th style=""width:189px;"">求職者検索</th>"
		sHTML2 = sHTML2 & "</tr>"
		sHTML2 = sHTML2 & "</thead>"
		sHTML2 = sHTML2 & "<tbody>"
		sHTML2 = sHTML2 & sHTML
		sHTML2 = sHTML2 & "</tbody>"
		sHTML2 = sHTML2 & "</table>" & vbCrLf
		sHTML2 = sHTML2 & sPageControl & vbCrLf
	End If

	GetHtmlMyMenuMyOrders = sHTML2
End Function

'******************************************************************************
'概　要：求人票登録権限可否チェック
'引　数：vOrderCode	：情報コード
'　　　：vUserID	：ログイン中ユーザコード
'　　　：vUseFlag	：ログイン中企業のライセンスの有効フラグ
'戻り値：Boolean	：[True]求人票登録可能 [False]求人票登録不可
'備　考：
'使用元：しごとナビ/company/order/edit1.asp
'更　新：2008/10/08 LIS K.kokubo 作成
'******************************************************************************
Function ChkEditOrder(ByVal vOrderCode, ByVal vUserID, ByVal vUseFlag)
	Dim sSQL
	Dim oRS
	Dim sError
	Dim flgQE

	Dim dbCheck
	Dim dbLicenseFlag

'	If vOrderCode = "" Then Exit Function

	'<ライセンス切れはマイメニューへリダイレクト>
	sSQL = "EXEC up_DtlNaviLicense_Now '" & vUserID & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		If oRS.Collect("LicenseType1Flag") <> "1" Then Exit Function
	End If
	Call RSClose(oRS)
	'</ライセンス切れはマイメニューへリダイレクト>

	'<ログイン中の企業の情報コードかどうかをチェック>
	sSQL = "sp_ChkCompanyOrder '" & vUserID & "', '" & vOrderCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		dbCheck = oRS.Collect("CheckFlag")
		dbLicenseFlag = oRS.Collect("LicenseFlag")
	End If
	Call RSClose(oRS)
	If vOrderCode = "" Then dbCheck = "1"
	If dbCheck = "0" And dbLicenseFlag = "0" Then Exit Function
	'</ログイン中の企業の情報コードかどうかをチェック>

	ChkEditOrder = True
End Function
%>
