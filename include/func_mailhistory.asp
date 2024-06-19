<%
'**********************************************************************************************************************
'概　要：メール一覧ページ	/staff/mailhistory_person_entity.asp
'　　　：上記ページで出力用の関数群をこのファイルに用意する。
'　　　：
'　　　：■■■　前提条件　■■■
'　　　：要事前インクルード
'　　　：/config/personel.asp
'　　　：/include/commonfunc.asp
'一　覧：■■■　メール一覧ページ出力用　■■■
'　　　：GetHtmlMailPageControlParam		：メール一覧ページのページコントロールのＨＴＭＬを取得
'　　　：GetHtmlMailHistory					：メール一覧ＨＴＭＬ取得
'　　　：GetHtmlMailHistoryDetailStaff		：スタッフのメール一覧の１行ＨＴＭＬを取得
'　　　：GetHtmlMailHistoryDetailCompany	：企業のメール一覧の１行ＨＴＭＬを取得
'　　　：GetHtmlMailControl					：削除ボタン、更新ボタン出力
'　　　：GetHtmlMailSearch					：メール検索出力
'　　　：DspNoMail							：メールが無い場合の出力
'　　　：
'　　　：■■■　メール管理　■■■
'　　　：UpdMailListStf	：メール一覧から評価と備考の更新
'　　　：DelMailStf		：メールの削除
'　　　：
'　　　：■■■　メール関連値取得　■■■
'　　　：GetSortImg		：並び替えのボタンを取得
'**********************************************************************************************************************

'******************************************************************************
'概　要：メール一覧ページのページコントロールのＨＴＭＬを取得
'引　数：rDB		：接続中ＤＢコネクション
'　　　：rRS		：up_SearchMail で生成されたレコードセットオブジェクト
'　　　：vUserType	：ログイン中ユーザ種類
'　　　：vMode		：現在の表示モード ["0"]受信メール一覧 ["1"]送信メール一覧
'　　　：vPageSize	：最大表示可能メール数
'　　　：vPage		：現在のページ
'　　　：vURL		：
'作成者：Lis Kokubo
'作成日：2007/03/02
'備　考：
'使用元：ナビ/staff/mailhistory_person_entity.asp
'******************************************************************************
Function GetHtmlMailPageControlParam(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vMode, ByVal vPageSize, ByVal vPage, ByVal vURL)
	Dim iPage
	Dim iPageSize
	Dim iMaxPage
	Dim iStartPage
	Dim iEndPage
	Dim idxPage
	Dim sAnc
	Dim sAncRcvSnd
	Dim sHtml

	GetHtmlMailPageControlParam = ""
	sHtml = ""

	If GetRSState(rRS) = False Then Exit Function

	If InStr(vURL, "?") > 0 Then
		sAnc = vURL & "&amp;"
	Else
		sAnc = vURL & "?"
	End If

	If vUserType = "staff" Then
		sAncRcvSnd = HTTP_CURRENTURL & "staff/mailhistory_person.asp"
	ElseIf vUserType = "company" Or vUserType = "dispatch" Then
		sAncRcvSnd = HTTP_CURRENTURL & "company/mailhistory_company.asp"
	End If

	iPage = CInt(vPage)
	iPageSize = CInt(vPageSize)

	rRS.PageSize = iPageSize
	iMaxPage = rRS.PageCount

	'範囲外のページ指定対策
	If iPage > iMaxPage Then iPage = iMaxPage
	If iPage < 1 Then iPage = 1

	rRS.AbsolutePage = iPage

	'表示開始ページ番号を指定
	iStartPage = iPage - 2
	If iStartPage < 1 Then iStartPage = 1

	'表示終了ページ番号を指定
	iEndPage = iPage + 2
	If iEndPage - iStartPage < 4 Then iEndPage = 5
	If iEndPage > iMaxPage Then iEndPage = iMaxPage

	sHtml = sHtml & "<div class=""mail_joken"">"

	If vMode = "1" Then
		sHtml = sHtml & "<h3>送信箱</h3>"
		sHtml = sHtml & "<div>"
		sHtml = sHtml & "<a href=""" & sAncRcvSnd & "?mode=0""><span>受信箱</span></a><span>送信箱</span>" & vbCrLf
		sHtml = sHtml & "<a href=""#mail_under""><span>送信メール検索</span></a></div>"
	Else
		sHtml = sHtml & "<h3>受信箱</h3>"
		sHtml = sHtml & "<div>"
		sHtml = sHtml & "<span>受信箱</span><a href=""" & sAncRcvSnd & "?mode=1""><span>送信箱</span></a>" & vbCrLf
		sHtml = sHtml & "<a href=""#mail_under""><span>受信メール検索</span></a></div>"
	End If

	sHtml = sHtml & "</div><br clear=""both"">" & vbCrLf
		sHtml = sHtml & "<div>" & vbCrLf
	sHtml = sHtml & "<div class=""left"">" & vbCrLf

	If CInt(iStartPage) <> 1 Then sHtml = sHtml & "…"
	For idxPage = iStartPage To iEndPage	'ページ番号を表示
		sHtml = sHtml & "　"
		If idxPage = CInt(iPage) Then		'指定ページの表示
			sHtml = sHtml & "[" & idxPage & "]"
		Else
			sHtml = sHtml & "<a href=""" & sAnc & "page=" & idxPage & """>" & idxPage & "</a>"
		End If
	Next
	If iEndPage < iMaxPage Then sHtml = sHtml & "　…"

	sHtml = sHtml & "</div>" & vbCrLf

	
	sHtml = sHtml & "<div class=""right"">" & rRS.RecordCount & "件ヒット：" & iPage & "/" & iMaxPage & "ページ目</div>" & vbCrLf
	sHtml = sHtml & "<div style=""clear:both;""></div>" & vbCrLf
	sHtml = sHtml & "</div>" & vbCrLf

	GetHtmlMailPageControlParam = sHtml
End Function


'******************************************************************************
'概　要：メール一覧ページの下のページコントロールのＨＴＭＬを取得

'******************************************************************************
Function GetHtmlMailPageControlParam2(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vMode, ByVal vPageSize, ByVal vPage, ByVal vURL)
	Dim iPage
	Dim iPageSize
	Dim iMaxPage
	Dim iStartPage
	Dim iEndPage
	Dim idxPage
	Dim sAnc
	Dim sAncRcvSnd
	Dim sHtml

	GetHtmlMailPageControlParam2 = ""
	sHtml = ""

	If GetRSState(rRS) = False Then Exit Function

	If InStr(vURL, "?") > 0 Then
		sAnc = vURL & "&amp;"
	Else
		sAnc = vURL & "?"
	End If

	If vUserType = "staff" Then
		sAncRcvSnd = HTTP_CURRENTURL & "staff/mailhistory_person.asp"
	ElseIf vUserType = "company" Or vUserType = "dispatch" Then
		sAncRcvSnd = HTTP_CURRENTURL & "company/mailhistory_company.asp"
	End If

	iPage = CInt(vPage)
	iPageSize = CInt(vPageSize)

	rRS.PageSize = iPageSize
	iMaxPage = rRS.PageCount

	'範囲外のページ指定対策
	If iPage > iMaxPage Then iPage = iMaxPage
	If iPage < 1 Then iPage = 1

	rRS.AbsolutePage = iPage

	'表示開始ページ番号を指定
	iStartPage = iPage - 2
	If iStartPage < 1 Then iStartPage = 1

	'表示終了ページ番号を指定
	iEndPage = iPage + 2
	If iEndPage - iStartPage < 4 Then iEndPage = 5
	If iEndPage > iMaxPage Then iEndPage = iMaxPage


		sHtml = sHtml & "<div>" & vbCrLf
	sHtml = sHtml & "<div class=""left"">" & vbCrLf

	If CInt(iStartPage) <> 1 Then sHtml = sHtml & "…"
	For idxPage = iStartPage To iEndPage	'ページ番号を表示
		sHtml = sHtml & "　"
		If idxPage = CInt(iPage) Then		'指定ページの表示
			sHtml = sHtml & "[" & idxPage & "]"
		Else
			sHtml = sHtml & "<a href=""" & sAnc & "page=" & idxPage & """>" & idxPage & "</a>"
		End If
	Next
	If iEndPage < iMaxPage Then sHtml = sHtml & "　…"

	sHtml = sHtml & "</div>" & vbCrLf

	
	sHtml = sHtml & "<div class=""right"">" & rRS.RecordCount & "件ヒット：" & iPage & "/" & iMaxPage & "ページ目</div>" & vbCrLf
	sHtml = sHtml & "<div style=""clear:both;""></div>" & vbCrLf
	sHtml = sHtml & "</div>" & vbCrLf

	GetHtmlMailPageControlParam2 = sHtml
End Function

'******************************************************************************
'概　要：メール一覧ページのメール一覧を出力
'引　数：rRS			：up_SearchMail で生成されたレコードセットオブジェクト
'　　　：vUserType		：
'　　　：vPage			：現在のページ
'　　　：vMode			：現在の表示モード ["0"]受信メール一覧 ["1"]送信メール一覧
'　　　：vSort			：並び替え ["0"]送信日降順 ["1"]送信日昇順
'　　　：vPageSize		：最大表示可能メール数
'　　　：vParamToDetail	：メール検索用パラメータ（メール詳細へ引き継ぐためのもの）
'　　　：vParamToSort	：ソート用パラメータ（page, sort パラメータを除いたもの）
'作成者：Lis Kokubo
'作成日：2007/03/02
'備　考：
'使用元：ナビ/staff/mailhistory_person_entity.asp
'******************************************************************************
Function GetHtmlMailHistory(ByRef rDB, ByRef rRS, ByRef vUserType, ByVal vPage, ByVal vMode, ByVal vSort, ByVal vPageSize, ByVal vParamToDetail, ByVal vParamToSort)
	Dim iLine				'出力中メール数
	Dim iPageSize			'１ページに出力するメール数の最大値
	Dim sEvaluationField	'評価 受信モードの場合：["ReceiverEvaluation"] 送信モードの場合：["SenderEvaluation"]
	Dim sRemarkField		'備考 受信モードの場合：["ReceiverRemark"] 送信モードの場合：["SenderRemark"]
	Dim sTableClass
	Dim sHTML

	If GetRSState(rRS) = False Then Exit Function

	If vUserType = "staff" Then
		sTableClass = "pattern1"
	ElseIf vUserType = "company" Or vUserType = "dispatch" Then
		sTableClass = "pattern8"
	End If

	'並び替え
	Select Case vSort
		Case "0": rRS.Sort = "SendDay DESC"
		Case "1": rRS.Sort = "SendDay ASC"
		Case "2": oRS.Sort = "OrderCode DESC"
		Case "3": oRS.Sort = "OrderCode ASC"
		Case "4": oRS.Sort = "ReceiverCode DESC"
		Case "5": oRS.Sort = "ReceiverCode ASC"
		Case "6": oRS.Sort = "SenderCode DESC"
		Case "7": oRS.Sort = "SenderCode ASC"
	End Select

	iPageSize = vPageSize
	rRS.PageSize = iPageSize

	'範囲外のページ指定対策
	If vPage > rRS.PageCount Then vPage = rRS.PageCount
	If vPage < 1 Then vPage = 1

	rRS.AbsolutePage = vPage

	sHTML = ""
	sHTML = sHTML & "<table class=""pattern1 mailHisTable smartNone"" border=""0"" cellspacing=""0"">"
	sHTML = sHTML & "<thead>"
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<th class=""mail_delete"">削除</th>"

	If vMode = "0" Then
		'受信箱画面
		sHTML = sHTML & "<th class=""mail_person"">差出人(情報コード)／件名</th>"
		sHTML = sHTML & "<th class=""mail_day"">受信日" & GetSortImg(vUserType, "0", vSort, vMode, vParamToSort) & "&nbsp;" & GetSortImg(vUserType, "1", vSort, vMode, vParamToSort) & "</th>"	
	Else
		'送信箱画面
		sHTML = sHTML & "<th class=""mail_person"">宛先(情報コード)／件名</th>"
		sHTML = sHTML & "<th class=""mail_day"">送信日" & GetSortImg(vUserType, "0", vSort, vMode, vParamToSort) & "&nbsp;" & GetSortImg(vUserType, "1", vSort, vMode, vParamToSort) & "</th>"
	End If
	
	sHTML = sHTML & "<th class=""mail_memo"">備考</th>"
	sHTML = sHTML & "<th class=""mail_point"">重要度</th>"
	sHTML = sHTML & "</tr>"
	sHTML = sHTML & "</thead>"
	
	
	sHTML = sHTML & "<tfoot>"
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<th class=""mail_delete"">削除</th>"

	If vMode = "0" Then
		'受信箱画面
		sHTML = sHTML & "<th class=""mail_person"">差出人(情報コード)／件名</th>"
		sHTML = sHTML & "<th class=""mail_day"">受信日" & GetSortImg(vUserType, "0", vSort, vMode, vParamToSort) & "&nbsp;" & GetSortImg(vUserType, "1", vSort, vMode, vParamToSort) & "</th>"	
	Else
		'送信箱画面
		sHTML = sHTML & "<th class=""mail_person"">宛先(情報コード)／件名</th>"
		sHTML = sHTML & "<th class=""mail_day"">送信日" & GetSortImg(vUserType, "0", vSort, vMode, vParamToSort) & "&nbsp;" & GetSortImg(vUserType, "1", vSort, vMode, vParamToSort) & "</th>"
	End If
	
	sHTML = sHTML & "<th class=""mail_memo"">備考</th>"
	sHTML = sHTML & "<th class=""mail_point"">重要度</th>"
	sHTML = sHTML & "</tr>"
	sHTML = sHTML & "</tfoot>"
	
	sHTML = sHTML & "<tbody>"

	iLine = 1
	Do While (GetRSState(rRS) = True And iLine <= vPageSize)
		If vUserType = "staff" Then sHTML = sHTML & GetHtmlMailHistoryDetailStaff(rDB, rRS, vUserType, vMode, vParamToDetail)
		If vUserType = "company" Or vUserType = "dispatch" Then sHTML = sHTML & GetHtmlMailHistoryDetailCompany(rDB, rRS, vUserType, vMode, vParamToDetail)

		iLine = iLine + 1
		rRS.MoveNext
	Loop
	

	sHTML = sHTML & "</tbody>"
	sHTML = sHTML & "</table>" & vbCrLf

	GetHtmlMailHistory = sHTML
End Function

'******************************************************************************
'概　要：スタッフのメール一覧ページの１行ＨＴＭＬを取得
'引　数：rDB		：接続中のＤＢオブジェクト
'　　　：rRS		：up_SearchMail で生成されたレコードセットオブジェクト
'　　　：vUserType	：ログイン中ユーザ種類
'　　　：vMode		：現在の表示モード ["0"]受信メール一覧 ["1"]送信メール一覧
'作成者：Lis Kokubo
'作成日：2007/03/02
'備　考：
'使用元：ナビ/staff/mailhistory_person_entity.asp
'******************************************************************************
Function GetHtmlMailHistoryDetailStaff(ByRef rDB, ByRef rRS, ByRef vUserType, ByVal vMode, ByVal vParam)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim iID
	Dim sAnswerNGFlag
	Dim sSendDay		'送受信日
	Dim sImgMailState	'メール開封・返信・未開封イメージ
	Dim sName			'対象の企業名・リス社員名
	Dim sOpenDay		'開封時間
	Dim sSubject		'メール件名
	Dim sAncOrder		'求人票へのリンク
	Dim sAncMailDetail	'メール詳細へのリンク
	Dim sEvaluation		'評価
	Dim sRemark			'備考
	Dim sHTML

	If GetRSState(rRS) = False Then Exit Function
	If vUserType <> "staff" Then Exit Function

	iID = rRS.Collect("ID")

	sSQL = "up_GetDetailMail '" & iID & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	If GetRSState(oRS) = False Then Exit Function

	'メール返信可否
	If oRS.Collect("SuspensionFlag") = "1" Or oRS.Collect("ErasureFlag") = "1" Then
		sAnswerNGFlag = "1"	'返信不可
	Else
		sAnswerNGFlag = "0"	'返信可
	End If

	'メール送受信日
	sSendDay = GetDateStr(oRS.Collect("SendDay"), "/") & "<br><span>" & GetTimeStr(oRS.Collect("SendDay"), ":") &"</span>"

	'受信モード時は、メールの状態画像を表示
	If vMode = "0" Then
		If oRS.Collect("AnswerFlag") = "1" Then
			'返信済
			sImgMailState = "<img src=""/img/common/mailre.gif"" alt=""返信済"">&nbsp;"
		ElseIf ChkStr(oRS.Collect("OpenDay")) <> "" Then
			'開封済
			sImgMailState = "<img src=""/img/common/mailkai.gif"" alt=""開封済"">&nbsp;"
		Else
			'未開封
			sImgMailState = "<img src=""/img/common/mailhei.gif"" alt=""未開封"">&nbsp;"
		End If
	End If

	'対象の企業名・リス社員名を表示
	If vMode = "1" Then
		'送信箱画面
		sName = oRS.Collect("ReceiverCompanyName")
	Else
		'受信箱画面
		sName = oRS.Collect("SenderCompanyName")
	End If

	'開封時間
	If ChkStr(oRS.Collect("OpenDay")) <> "" Then
		sOpenDay = "<span class=""moji812"">【開封：" & ChkStr(oRS.Collect("OpenDay")) & "】</span><br>"
	End If

	'メール件名
	If oRS.Collect("Subject") <> "" Then
		sSubject = oRS.Collect("Subject")
	Else
		sSubject = "タイトルなし"
	End If

	'求人票へのリンク
	If Left(oRS.Collect("OrderCode"),1) = "J" Then
		sAncOrder = sAncOrder & "&nbsp;("
		If ChkOrderDsp(rDB, oRS.Collect("OrderCode"), G_USERID) = True Then
			'掲載中の求人票の場合は求人票へのリンク
			sAncOrder = sAncOrder & "<a href=""" & HTTP_CURRENTURL & "order/order_detail.asp?OrderCode=" & oRS.Collect("OrderCode") & """>" & oRS.Collect("OrderCode") & "</a>"
		Else
			'非掲載の求人票の場合は情報コードのみ
			sAncOrder = sAncOrder & oRS.Collect("OrderCode")
		End If
		If oRS.Collect("SecretFlag") = "1" Then sAncOrder = sAncOrder & "&nbsp;<img src=""/img/order/secret.gif"" alt=""スカウトを受けた人だけが閲覧できる求人情報"" border=""0"">"
		sAncOrder = sAncOrder & ")"
		sAncOrder = sAncOrder & "<br>"
	Else
		sAncOrder = "<br>"
	End If

	'評価・備考
	If vMode = "1" Then
		sEvaluation = ChkStr(oRS.Collect("SenderEvaluation"))
		sRemark = ChkStr(oRS.Collect("SenderRemark"))
	Else
		sEvaluation = ChkStr(oRS.Collect("ReceiverEvaluation"))
		sRemark = ChkStr(oRS.Collect("ReceiverRemark"))
	End If

	sAncMailDetail = HTTPS_CURRENTURL & "staff/mail_detail_person.asp"
	If vParam <> "" Then
		sAncMailDetail = sAncMailDetail & vParam & "&amp;id=" & iID
	Else
		sAncMailDetail = sAncMailDetail & "?id=" & iID
	End If

	sHTML = ""
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<th class=""delCheck""><input type=""checkbox"" name=""delflag"" value=""" & iID & """></th>"
	sHTML = sHTML & "<td class=""fromWho"">" & sImgMailState & sName &  sAncOrder & sOpenDay & "<a href=""" & sAncMailDetail & """>" & sSubject & "</a></td>"
	sHTML = sHTML & "<td class=""mailDay"">" & sSendDay & "</td>"
	sHTML = sHTML & "<td class=""textareaMemo"">"
	sHTML = sHTML & "<input type=""text"" name=""CONF_Remark" & iID & """ value=""" & sRemark & """>"
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "<td class=""weightCheck"">"
	sHTML = sHTML & "<select name=""CONF_Evaluation" & iID & """>"
	sHTML = sHTML & "<option value=""""></option>"
	If sEvaluation = "A" Then: sHTML = sHTML & "<option value=""A"" selected>Ａ</option>": Else: sHTML = sHTML & "<option value=""A"">Ａ</option>": End If
	If sEvaluation = "B" Then: sHTML = sHTML & "<option value=""B"" selected>Ｂ</option>": Else: sHTML = sHTML & "<option value=""B"">Ｂ</option>": End If
	If sEvaluation = "C" Then: sHTML = sHTML & "<option value=""C"" selected>Ｃ</option>": Else: sHTML = sHTML & "<option value=""C"">Ｃ</option>": End If
	If sEvaluation = "D" Then: sHTML = sHTML & "<option value=""D"" selected>Ｄ</option>": Else: sHTML = sHTML & "<option value=""D"">Ｄ</option>": End If
	If sEvaluation = "E" Then: sHTML = sHTML & "<option value=""E"" selected>Ｅ</option>": Else: sHTML = sHTML & "<option value=""E"">Ｅ</option>": End If
	sHTML = sHTML & "</select>"
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"

	GetHtmlMailHistoryDetailStaff = sHTML
End Function

'******************************************************************************
'概　要：企業のメール一覧ページの１行ＨＴＭＬを取得
'引　数：rDB		：接続中のＤＢオブジェクト
'　　　：rRS		：up_SearchMail で生成されたレコードセットオブジェクト
'　　　：vUserType	：ログイン中ユーザ種類
'　　　：vMode		：現在の表示モード ["0"]受信メール一覧 ["1"]送信メール一覧
'　　　：vParam		：メール検索パラメータ
'使用元：ナビ/staff/mailhistory_person_entity.asp
'備　考：
'更　新：2007/03/02 LIS K.Kokubo 作成
'　　　：2008/09/09 LIS K.Kokubo 求職者名入力対応
'******************************************************************************
Function GetHtmlMailHistoryDetailCompany(ByRef rDB, ByRef rRS, ByRef vUserType, ByVal vMode, ByVal vParam)
	Dim sSQL
	Dim oRS
	Dim oRS2
	Dim flgQE
	Dim sError

	Dim dbStaffName

	Dim iID
	Dim sStaffCode
	Dim sAnswerNGFlag
	Dim sSendDay		'送受信日
	Dim sImgMailState	'メール開封・返信・未開封イメージ
	Dim sName			'対象の企業名・リス社員名
	Dim sOpenDay		'開封時間
	Dim sSubject		'メール件名
	Dim sAncOrder		'求人票へのリンク
	Dim sAncStaff		'求職者プロフィールへのリンク
	Dim sAncProgress	'進捗状況へのリンク
	Dim sAncMailDetail	'メール詳細へのリンク
	Dim sEvaluation		'評価
	Dim sRemark			'備考
	Dim sHTML

	If GetRSState(rRS) = False Then Exit Function
	If Not(vUserType = "company" Or vUserType = "dispatch") Then Exit Function

	iID = rRS.Collect("ID")

	sSQL = "up_GetDetailMail '" & iID & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	'求職者コードの取得
	If vMode = "1" Then
		'送信モードの場合は受信者コードが求職者コード
		sStaffCode = oRS.Collect("ReceiverCode")
	Else
		'受信モードの場合は送信者コードが求職者コード
		sStaffCode = oRS.Collect("SenderCode")
	End If

	'メール返信可否
	If oRS.Collect("SuspensionFlag") = "1" Or oRS.Collect("ErasureFlag") = "1" Then
		sAnswerNGFlag = "1"	'返信不可
	Else
		sAnswerNGFlag = "0"	'返信可
	End If

	'メール送受信日
	sSendDay = GetDateStr(oRS.Collect("SendDay"), "/") & "<br>" & GetTimeStr(oRS.Collect("SendDay"), ":")

	'受信モード時は、メールの状態画像を表示
	If vMode = "0" Then
		If oRS.Collect("AnswerFlag") = "1" Then
			'返信済
			sImgMailState = "<img src=""/img/common/mailre.gif"" alt=""返信済"">&nbsp;"
		ElseIf ChkStr(oRS.Collect("OpenDay")) <> "" Then
			'開封済
			sImgMailState = "<img src=""/img/common/mailkai.gif"" alt=""開封済"">&nbsp;"
		Else
			'未開封
			sImgMailState = "<img src=""/img/common/mailhei.gif"" alt=""未開封"">&nbsp;"
		End If
	End If

	'対象の企業名・リス社員名を表示
	If vMode = "1" Then
		'送信箱画面
		sName = oRS.Collect("ReceiverCompanyName")
	Else
		'受信箱画面
		sName = oRS.Collect("SenderCompanyName")
	End If

	'求人票へのリンク
	sAncOrder = "(<a href=""/order/order_detail.asp?OrderCode=" & oRS.Collect("OrderCode") & """>" & oRS.Collect("OrderCode") & "</a>)"

	'開封時間
	If ChkStr(oRS.Collect("OpenDay")) <> "" Then
		sOpenDay = "<span class=""moji812"">【開封：" & ChkStr(oRS.Collect("OpenDay")) & "】</span><br>"
	End If

	'メール件名
	If oRS.Collect("Subject") <> "" Then
		sSubject = oRS.Collect("Subject")
	Else
		sSubject = "タイトルなし"
	End If

	'求職者プロフィールへのリンク
	If Left(sStaffCode,1) = "S" Then
		'<求職者名取得>
		sSQL = "EXEC up_DtlCMPStaffName '" & G_USERID & "', '" & sStaffCode & "'"
		flgQE = QUERYEXE(dbconn, oRS2, sSQL, sError)
		If GetRSState(oRS2) = True Then
			dbStaffName = ChkStr(oRS2.Collect("StaffName"))
		End If
		Call RSClose(oRS2)
		'</求職者名取得>

		If sAnswerNGFlag = "0" Then
			'返信可能
			sAncStaff = "/staff/person_detail.asp?staffcode=" & sStaffCode
			If oRS.Collect("OrderPublicFlag") = "1" Then sAncStaff = sAncStaff & "&amp;ordercode=" & oRS.Collect("OrderCode")
			sAncStaff = "<a href=""" & sAncStaff & """>"
			If dbStaffName <> "" Then
				sAncStaff = sAncStaff & dbStaffName
			Else
				sAncStaff = sAncStaff & sStaffCode
			End If
			sAncStaff = sAncStaff & "</a>"
		Else
			'返信不可の場合は情報コードのみ
			If dbStaffName <> "" Then
				sAncStaff = dbStaffName
			Else
				sAncStaff = sStaffCode
			End If
		End If

		If vParam <> "" Then
			sAncProgress = "<a href=""" & HTTPS_CURRENTURL & "company/mailhistory_progress.asp" & vParam & "&amp;staffcode=" & sStaffCode & "&amp;ordercode=" & oRS.Collect("OrderCode") & """>進捗確認</a><br>"
		Else
			sAncProgress = "<a href=""" & HTTPS_CURRENTURL & "company/mailhistory_progress.asp?staffcode=" & sStaffCode & "&amp;ordercode=" & oRS.Collect("OrderCode") & """>進捗確認</a><br>"
		End If
	Else
		'企業コラボ
		'Response.Write "<a href=""../ad/ad_detail.asp?advertcode=" & oRS.Collect("OrderCode") & """>" & oRS.Collect("OrderCode") & "</a>"
	End If

	'評価・備考
	If vMode = "1" Then
		sEvaluation = ChkStr(oRS.Collect("SenderEvaluation"))
		sRemark = ChkStr(oRS.Collect("SenderRemark"))
	Else
		sEvaluation = ChkStr(oRS.Collect("ReceiverEvaluation"))
		sRemark = ChkStr(oRS.Collect("ReceiverRemark"))
	End If

	'メール詳細へのリンク
	sAncMailDetail = HTTPS_CURRENTURL & "company/mail_detail_company.asp"
	If vParam <> "" Then
		sAncMailDetail = sAncMailDetail & vParam & "&amp;id=" & iID
	Else
		sAncMailDetail = sAncMailDetail & "?id=" & iID
	End If

	sHTML = ""
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<th><input type=""checkbox"" name=""delflag"" value=""" & iID & """></th>"
	sHTML = sHTML & "<td>" & sImgMailState & sAncStaff & "&nbsp;" & sAncOrder & "&nbsp;-&nbsp;" & sAncProgress & "<a href=""" & sAncMailDetail & """>" & sSubject & "</a></td>"
	sHTML = sHTML & "<td>" & sSendDay & "</td>"
	sHTML = sHTML & "<td>"
	sHTML = sHTML & "<input type=""text"" name=""CONF_Remark" & iID & """ value=""" & sRemark & """ size=""20"">"
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "<td>"
	sHTML = sHTML & "<select name=""CONF_Evaluation" & iID & """>"
	sHTML = sHTML & "<option value=""""></option>"
	If sEvaluation = "A" Then: sHTML = sHTML & "<option value=""A"" selected>Ａ</option>": Else: sHTML = sHTML & "<option value=""A"">Ａ</option>": End If
	If sEvaluation = "B" Then: sHTML = sHTML & "<option value=""B"" selected>Ｂ</option>": Else: sHTML = sHTML & "<option value=""B"">Ｂ</option>": End If
	If sEvaluation = "C" Then: sHTML = sHTML & "<option value=""C"" selected>Ｃ</option>": Else: sHTML = sHTML & "<option value=""C"">Ｃ</option>": End If
	If sEvaluation = "D" Then: sHTML = sHTML & "<option value=""D"" selected>Ｄ</option>": Else: sHTML = sHTML & "<option value=""D"">Ｄ</option>": End If
	If sEvaluation = "E" Then: sHTML = sHTML & "<option value=""E"" selected>Ｅ</option>": Else: sHTML = sHTML & "<option value=""E"">Ｅ</option>": End If
	sHTML = sHTML & "</select>"
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"

	GetHtmlMailHistoryDetailCompany = sHTML
End Function

'******************************************************************************
'概　要：メール一覧ページの管理ボタンを出力
'引　数：vMode		：現在の表示モード ["0"]受信メール一覧 ["1"]送信メール一覧
'作成者：Lis Kokubo
'作成日：2007/03/02
'備　考：
'使用元：ナビ/staff/mailhistory_person_entity.asp
'******************************************************************************
Function GetHtmlMailControl(ByVal vMode)
	Dim sHTML
	Dim sOnClickDel	'削除ボタンをクリックした際のjavascript
	Dim sOnClickUpd	'更新ボタンをクリックした際のjavascript

	sOnClickDel = "if(ChkInput(getElementsByName('delflag'), 'checkbox', '1', '削除するメールにチェックしてください。') == true){if(confirm('チェックしたメールを削除しますか？') == true){this.form.hdndelflag.value='1';this.form.submit();}}"
	sOnClickUpd = "this.form.hdnupdflag.value='1';this.form.submit();"

	sHTML = ""
	sHTML = sHTML & "<br>"
	sHTML = sHTML & "<div id=""mail_input"">"
	sHTML = sHTML & "<input type=""button"" name=""DeleteData"" value=""選択したメールを削除"" onclick=""" & sOnClickDel & """>"
	sHTML = sHTML & "<input id=""hdndelflag"" type=""hidden"" name=""frmdelflag"" value="""">"
	sHTML = sHTML & "<div>"
	sHTML = sHTML & "<input type=""button"" name=""Update"" value=""備考・重要度を更新"" onclick=""" & sOnClickUpd & """>"
	sHTML = sHTML & "<input id=""hdnupdflag"" type=""hidden"" name=""frmupdflag"" value="""">"
	sHTML = sHTML & "<br><span>※「備考」「重要度」はご自身のメモ機能<br>としてお使いいただけます。</span>"
	sHTML = sHTML & "</div></div>"
	sHTML = sHTML & "<br clear=""both"">"
	

	GetHtmlMailControl = sHTML
End Function

'******************************************************************************
'概　要：メール一覧ページのメール検索を出力
'引　数：vUserID	：ユーザＩＤ
'　　　：vUserType	：ログイン中ユーザ種類
'　　　：vMode		：現在の表示モード ["0"]受信メール一覧 ["1"]送信メール一覧
'　　　：vSort		：現在の並び替えフラグ
'　　　：rSMC		：clsSearchMailConditionのインスタンス
'使用元：ナビ/staff/mailhistory_person_entity.asp
'　　　：ナビ/company/mailhistory_company.asp
'備　考：
'履　歴：2007/03/02 LIS K.Kokubo 作成
'　　　：2009/07/30 LIS K.Kokubo スカウトアプローチ検索追加
'******************************************************************************
Function GetHtmlMailSearch(ByVal vUserID, ByVal vUserType, ByVal vMode, ByVal vSort, ByRef rSMC)
	Dim sHTML
	Dim sTableClass
	Dim sChecked

	If vUserType = "staff" Then
		sTableClass = "pattern1"
	ElseIf vUserType = "company" Or vUserType = "dispatch" Then
		sTableClass = "pattern3"
	End If

	sHTML = ""
	sHTML = sHTML & "<div id=""mail_search"" class=""smartNone"">"
	sHTML = sHTML & "<form id=""frmsearchmail"" action="""" method=""get"">"
	sHTML = sHTML & "<input type=""hidden"" name=""mode"" value=""" & vMode & """>"
	sHTML = sHTML & "<input type=""hidden"" name=""sort"" value=""" & vSort & """>"


	If vMode = "1" Then
	%>
	<h3 id="mail_under" class="smartNone">送信メール検索</h3>
    <%
	ElseIf vMode = "0" Then
	%>
	<h3 id="mail_under" class="smartNone">受信メール検索</h3>
    <%
    End If



	If vUserType = "company" Or vUserType = "dispatch" Then


		sHTML = sHTML & "<p>求人担当者</p>"
		sHTML = sHTML & "<select name=""mcpn"">"
		sHTML = sHTML & "<option value="""">--担当者--</option>"
		sHTML = sHTML & GetMailContactPersonNameOptionHtml(vUserID, vMode, rSMC.MailContactPersonName)
		sHTML = sHTML & "</select>"

		'<スカウト,アプローチ>

		If vMode = "1" Then
			sHTML = sHTML & "<pスカウトメール</p>"
		Else
			sHTML = sHTML & "<p>アプローチメール</p>"
		End If


		sChecked = ""
		If rSMC.ScoutApproachFlag = "1" Then sChecked = " checked"
		sHTML = sHTML & "<label><input name=""ssaf"" type=""radio"" value=""1""" & sChecked & ">"
		If vMode = "1" Then
			sHTML = sHTML & "スカウトメールのみ"
		Else
			sHTML = sHTML & "アプローチメールのみ"
		End If
		sHTML = sHTML & "</label>&nbsp;"

		sChecked = ""
		If rSMC.ScoutApproachFlag = "0" Then sChecked = " checked"
		sHTML = sHTML & "<label><input name=""ssaf"" type=""radio"" value=""0""" & sChecked & ">"
		If vMode = "1" Then
			sHTML = sHTML & "スカウトメール以外"
		Else
			sHTML = sHTML & "アプローチメール以外"
		End If
		sHTML = sHTML & "</label>&nbsp;"

		sChecked = ""
		If rSMC.ScoutApproachFlag = "" Then sChecked = " checked"
		sHTML = sHTML & "<label><input name=""ssaf"" type=""radio"" value=""""" & sChecked & ">指定しない</label>"

		'</スカウト,アプローチ>

		'<未読メール>
		If vMode = "0" Then

			sHTML = sHTML & "<p>開封状況</p>"


			sChecked = ""
			If rSMC.NotOpenFlag = "1" Then sChecked = " checked"
			sHTML = sHTML & "<label><input type=""radio"" name=""snof"" value=""1""" & sChecked & ">未読メールのみ</label>&nbsp;"

			sChecked = ""
			If rSMC.NotOpenFlag = "0" Then sChecked = " checked"
			sHTML = sHTML & "<label><input type=""radio"" name=""snof"" value=""0""" & sChecked & ">既読メールのみ</label>&nbsp;"

			sChecked = ""
			If Not(rSMC.NotOpenFlag = "1" Or rSMC.NotOpenFlag = "0") Then sChecked = " checked"
			sHTML = sHTML & "<label><input type=""radio"" name=""snof"" value=""""" & sChecked & ">指定しない</label>"


		End If
		'</未読メール>
	End If


	sHTML = sHTML & "<p><b>日付</b>で検索</p>"
	sHTML = sHTML & "<input class=""num8"" type=""text"" name=""sdf"" maxlength=""8"" value=""" & rSMC.DayFrom & """>"
	sHTML = sHTML & "～"
	sHTML = sHTML & "<input class=""num8"" type=""text"" name=""sdt"" maxlength=""8"" value=""" & rSMC.DayTo & """>（例：20020101）"
	sHTML = sHTML & "<p><b>情報コード</b>で検索</p>"
	sHTML = sHTML & "<input class=""alpha8"" type=""text"" name=""soc"" value=""" & rSMC.OrderCode & """ size=""15"">&nbsp;&nbsp;"
	
	If vUserType = "company" Or vUserType = "dispatch" Then
		If vMode = "1" Then
			sHTML = sHTML & "受信者コード"
		Else
			sHTML = sHTML & "送信者コード"
		End If
		sHTML = sHTML & "：<input class=""alpha8"" type=""text"" name=""sc"" value=""" & rSMC.SearchCode & """>"
	End If
	
	sHTML = sHTML & "<p><b>評価,件名,内容,備考</b>で検索</p>"
	sHTML = sHTML & "<select name=""se"">"
	sHTML = sHTML & "<option value="""">--評価--</option>"
	If rSMC.Evaluation = "A" Then: sHTML = sHTML & "<option value=""A"" selected>Ａ</option>": Else: sHTML = sHTML & "<option value=""A"">Ａ</option>": End If
	If rSMC.Evaluation = "B" Then: sHTML = sHTML & "<option value=""B"" selected>Ｂ</option>": Else: sHTML = sHTML & "<option value=""B"">Ｂ</option>": End If
	If rSMC.Evaluation = "C" Then: sHTML = sHTML & "<option value=""C"" selected>Ｃ</option>": Else: sHTML = sHTML & "<option value=""C"">Ｃ</option>": End If
	If rSMC.Evaluation = "D" Then: sHTML = sHTML & "<option value=""D"" selected>Ｄ</option>": Else: sHTML = sHTML & "<option value=""D"">Ｄ</option>": End If
	If rSMC.Evaluation = "E" Then: sHTML = sHTML & "<option value=""E"" selected>Ｅ</option>": Else: sHTML = sHTML & "<option value=""E"">Ｅ</option>": End If
	sHTML = sHTML & "</select>"
	sHTML = sHTML & "<input type=""text"" name=""skwd"" value=""" & rSMC.Keyword & """ maxlength=""50"" style=""width:300px;"">"
	sHTML = sHTML & "<br>"
	sHTML = sHTML & "<div align=""center""><input type=""submit"" value=""この条件で検索する""></div>"
	sHTML = sHTML & "</form>"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & vbCrLf

	GetHtmlMailSearch = sHTML
End Function

'******************************************************************************
'概　要：メールがない場合の出力
'引　数：vMode		：現在の表示モード ["0"]受信メール一覧 ["1"]送信メール一覧
'作成者：Lis Kokubo
'作成日：2007/03/02
'備　考：
'使用元：ナビ/staff/mailhistory_person_entity.asp
'******************************************************************************
Function GetHtmlNoMail(ByVal vUserType, ByVal vMode)
	Dim sHTML
	Dim sAnc

	If vUserType = "staff" Then
		sAnc = "./mailhistory_person.asp"
	ElseIf vUserType = "company" Or vUserType = "dispatch" Then
		sAnc = "./mailhistory_company.asp"
	End If

	sHTML = ""
	If vMode = "1" Then
		sHTML = sHTML & "<div style=""width:100%; text-align:center;"">"
		sHTML = sHTML & "<a href=""" & sAnc & "?mode=0"">受信箱</a>&nbsp;&nbsp;送信箱<br>"
		sHTML = sHTML & "送信メールはありません"
		sHTML = sHTML & "</div>"
		sHTML = sHTML & vbCrLf
	Else
		sHTML = sHTML & "<div style=""width:100%; text-align:center;"">"
		sHTML = sHTML & "受信箱&nbsp;&nbsp;<a href=""" & sAnc & "?mode=1"">送信箱</a><br>"
		sHTML = sHTML & "受信メールはありません"
		sHTML = sHTML & "</div>"
		sHTML = sHTML & vbCrLf
	End If

	GetHtmlNoMail = sHTML
End Function

'******************************************************************************
'概　要：メール一覧から評価と備考の更新
'引　数：vUserID	：ログイン中のユーザID
'　　　：vMode		：現在の表示モード ["0"]受信メール一覧 ["1"]送信メール一覧
'作成者：Lis Kokubo
'作成日：2007/03/02
'備　考：
'使用元：ナビ/staff/mailhistory_person_entity.asp
'******************************************************************************
Function UpdMailList(ByVal vUserID, ByVal vMode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sEvaluation
	Dim sRemark
	Dim sID
	Dim idx

	sSQL = ""
	For idx = 1 To Request.Form.Count
		sID = Mid(Request.Form.Key(idx), Len("CONF_Remark") + 1)
		If IsNumber(sID, 0, False) = True Then
			sEvaluation = GetForm("CONF_Evaluation" & sID, 1)
			sRemark = GetForm("CONF_Remark" & sID, 1)
			sSQL = sSQL & _
				"EXEC sp_Reg_MailUpdate '" & sID & "', '" & vUserID & "', '" & vMode & "', '" & sEvaluation & "', '" & sRemark & "'" & vbCrLf
		End If
	Next
	If sSQL <> "" Then flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
End Function

'******************************************************************************
'概　要：メールの削除
'引　数：vUserID	：ログイン中のユーザID
'　　　：vDelFlag	：削除対象メールID　[XXXX,XXXXX,XXXXXX,…]
'　　　：vMode		：現在の表示モード ["0"]受信メール一覧 ["1"]送信メール一覧
'作成者：Lis Kokubo
'作成日：2007/03/02
'備　考：
'使用元：ナビ/staff/mailhistory_person_entity.asp
'******************************************************************************
Function DelMailList(ByVal vUserID, ByVal vDelFlag, ByVal vMode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim aID
	Dim idx

	aID = Split(vDelFlag, ",")

	sSQL = ""
	For idx = 0 To UBound(aID)
		If IsNumber(aID(idx), 0, False) = False Then sSQL = "": Exit For
		sSQL = sSQL & "EXEC sp_Reg_MailDeleteFlag '" & aID(idx) & "', '" & vUserID & "', '" & vMode & "'" & vbCrLf
	Next
	If sSQL <> "" Then flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
End Function

'******************************************************************************
'概　要：並び替えのボタンを取得
'引　数：vUserType	：
'　　　：vMySortNo	：ボタン自身に割り当てられているソート値
'　　　：vSortNo	：並び替え ["0"]送信日降順 ["1"]送信日昇順
'　　　：vMode		：現在の表示モード ["0"]受信メール一覧 ["1"]送信メール一覧
'　　　：vParam		：メール検索条件パラメータ
'作成者：Lis Kokubo
'作成日：2007/03/02
'備　考：
'使用元：ナビ/staff/mailhistory_person_entity.asp
'******************************************************************************
Function GetSortImg(ByVal vUserType, ByVal vMySortNo, ByVal vSortNo, ByVal vMode, ByVal vParam)
	Dim sImg		'並び替え用ボタンのイメージファイル
	Dim sAlt		'並び替えようボタンのALT
	Dim sOnClick	'並び替え用ボタンをクリックしたときのjavascript
	Dim sURL

	'******************************************************************************
	'並び替え用の設定
	'------------------------------------------------------------------------------
	sImg = "/img/common/sort"

	If vMySortNo = vSortNo Then
		'現在の並び替え順が、自信の並び替え順の場合のイメージ
		sImg = sImg & "1"
	Else
		sImg = sImg & "4"
	End If

	If vUserType = "company" Or vUserType = "dispatch" Then
		sURL = HTTP_CURRENTURL & "company/mailhistory_company.asp"
	ElseIf vUserType = "staff" Then
		sURL = HTTP_CURRENTURL & "staff/mailhistory_person.asp"
	End If

	If CInt(vMySortNo) Mod 2 = 0 Then
		'降順
		sImg = sImg & "1"
		sAlt = "降順"
		If vParam <> "" Then
			sURL = sURL & vParam & "&amp;sort=0"
		Else
			sURL = sURL & "?sort=0"
		End If
	Else
		'昇順
		sImg = sImg & "2"
		sAlt = "昇順"
		If vParam <> "" Then
			sURL = sURL & vParam & "&amp;sort=1"
		Else
			sURL = sURL & "?sort=1"
		End If
	End If

	sImg = sImg & ".gif"
	'------------------------------------------------------------------------------
	'並び替え用の設定
	'******************************************************************************

	If vMySortNo = vSortNo Then
		'現在の並び替え順が、自信の並び替え順の場合のイメージ
		GetSortImg = "<img src=""" & sImg & """ alt=""" & sAlt & """ border=""0"">"
	Else
		GetSortImg = "<a href=""" & sURL & """><img src=""" & sImg & """ alt=""" & sAlt & """ border=""0""></a>"
	End If
End Function
%>
