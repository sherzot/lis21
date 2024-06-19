<%
'**********************************************************************************************************************
'概　要：メールテンプレート管理画面 /mailtemplate/manager.asp
'　　　：プロフィール /staff/person_detail.asp
'　　　：上記ページで出力用の関数群をこのファイルに用意する。
'　　　：
'　　　：■■■　前提条件　■■■
'　　　：要事前インクルード
'　　　：/config/personel.asp
'　　　：/include/commonfunc.asp
'一　覧：■■■　メールテンプレート管理画面出力用　■■■
'　　　：DspMailTemplateList				：メールテンプレート一覧部分出力
'　　　：DspMailTemplateListOne				：メールテンプレート一覧の個々のテンプレートを出力
'　　　：GetMailTemplateLink				：対象求人情報の全メールテンプレートの編集ページへのリンクを取得
'　　　：
'　　　：■■■　メールテンプレート入力・確認・登録完了画面出力用　■■■
'　　　：DspInputMailTemplate				：メールテンプレートの入力画面出力
'　　　：DspConfMailTemplate				：メールテンプレートの入力内容確認画面出力
'　　　：DspRegMailTemplate					：メールテンプレートＤＢ登録～登録完了・失敗画面出力
'　　　：
'　　　：■■■　メールテンプレート参照画面出力用　■■■
'　　　：DspMailTemplateRefList				：メールテンプレート参照画面のテンプレート一覧を出力
'　　　：DspMailTemplateRefListOne			：メールテンプレート参照画面の一覧の個々のテンプレートを出力
'　　　：GetMailTemplateRefLink				：メールテンプレート参照画面の個々のテンプレートの内容確認のページへのリンクを取得
'　　　：DspMailTemplateRefCopy				：メールテンプレート参照画面のテンプレート詳細＆コピーボタンを出力
'　　　：GetContactPersonNameMTOptionHtml	：メールテンプレートを保持する求人票の担当者一覧を <option></option> 形式で取得
'　　　：
'　　　：■■■　メールテンプレート削除画面出力用　■■■
'　　　：DspConfDeleteMailTemplate			：メールテンプレートの削除確認画面出力
'　　　：DspDeleteMailTemplate				：メールテンプレートＤＢ削除～削除完了画面出力
'　　　：
'　　　：■■■　メールテンプレートのコピー作成画面出力用　■■■
'　　　：DspCopyMailTemplate				：メールテンプレートコピー画面のコピー元を出力
'　　　：DspCopyMailTemplateList			：メールテンプレートコピー画面の一覧部分出力
'　　　：DspMailTemplateListOne2			：メールテンプレートコピー画面のコピー先求人票一覧を出力
'　　　：DspRegCopyMailTemplate				：メールテンプレートコピーＤＢ登録～登録完了画面出力
'　　　：
'　　　：■■■　メールテンプレートＤＢ処理　■■■
'　　　：RegMailTemplate					：メールテンプレートの登録処理
'　　　：DelMailTemplate					：メールテンプレートの削除処理
'**********************************************************************************************************************

'******************************************************************************
'概　要：メールテンプレート一覧部分出力
'引　数：rDB				：接続中のDBConnection
'　　　：vUserCode			：ログイン中ユーザ
'　　　：vContactPersonName	：求人担当者フィルタ
'　　　：vPageSize			：１ページあたりの表示件数
'　　　：vPage				：表示中ページ
'作成者：Lis Kokubo
'作成日：2007/06/15
'備　考：
'使用元：しごとナビ/mailtemplate/manager.asp
'******************************************************************************
Function DspMailTemplateList(ByRef rDB, ByVal vUserCode, ByVal vContactPersonName, ByVal vPageSize, ByVal vPage)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderCode
	Dim sHtmlPageCtrl
	Dim iRow
	Dim sFilterPerson	'求人担当者一覧（<option></option> を保持する）

	sSQL = "up_GetMyOrder '" & vUserCode & "', '0', '', ''"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	'求人担当者一覧を取得
	If GetRSState(oRS) = True Then
		sFilterPerson = GetContactPersonNameOptionHtml(G_USERID, vContactPersonName)
	End If

	'求人担当者で絞込み
	If GetRSState(oRS) = True And vContactPersonName <> "" Then
		If vContactPersonName <> "" Then oRS.Filter = "ContactPersonName = '" & vContactPersonName & "'"
	End If

	'ページコントロール取得
	If GetRSState(oRS) = True Then
		sHtmlPageCtrl = GetPageControlHtml(rDB, oRS, vPageSize, vPage)
	End If

	Response.Write sHtmlPageCtrl
%>
	<table class="pattern8" border="0">
		<thead>
			<tr>
				<th colspan="3">メールテンプレート一覧</th>
			</tr>
		</thead>
		<tbody>
			<tr>
				<th style="width:68px; text-align:center;">情報コード</th>
				<th style="width:109px;">
					<select id="contactpersonname" name="frmcontactpersonname" style="width:109px;" onchange="ChgPage(1);">
						<option value="">(求人担当者)</otpion>
						<%= sFilterPerson %>
					</select>
				</th>
				<th style="width:389px;">メールテンプレート種類</th>
			</tr>
<%
	iRow = 1
	If GetRSState(oRS) = True Then
		Do While GetRSState(oRS) = True And iRow <= vPageSize
			Call DspMailTemplateListOne(rDB, oRS, G_USERID)
			oRS.MoveNext
			iRow = iRow + 1
		Loop
	Else
%>
			<tr>
				<td colspan="3">求人票がありません。</td>
			</tr>
<%
	End If
%>
		</tbody>
	</table>
<%
	Response.Write sHtmlPageCtrl
	Call RSClose(oRS)
End Function

'******************************************************************************
'概　要：メールテンプレート一覧の個々のテンプレートを出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_GetMyOrderで生成されたレコードセットオブジェクト
'作成者：Lis Kokubo
'作成日：2007/06/15
'備　考：
'使用元：しごとナビ/mailtemplate/manager.asp
'******************************************************************************
Function DspMailTemplateListOne(ByRef rDB, ByRef rRS, ByVal vUserCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderCode
	Dim sContactPersonName
	Dim iMTCnt
	Dim sAnc				'求人票へのリンク

	If GetRSState(rRS) = False Then Exit Function

	sOrderCode = ChkStr(rRS.Collect("OrderCode"))
	sContactPersonName = ChkStr(rRS.Collect("ContactPersonName"))

	iMTCnt = 0
	sSQL = "up_ChkMyMailTemplate '" & vUserCode & "', '" & sOrderCode & "', ''"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		iMTCnt = oRS.Collect("MailTemplateCnt")
	End If

	sAnc = "<a href=""" & HTTPS_NAVI_CURRENTURL & "order/order_detail.asp?ordercode=" & sOrderCode & """>" & sOrderCode & "</a>"
%>
			<tr>
				<td style="text-align:center;"><%= sAnc %></td>
				<td><%= sContactPersonName %></td>
				<td>
					<%= GetMailTemplateLink(rDB, vUserCode, sOrderCode, iMTCnt) %>
				</td>
			</tr>
<%
End Function

'******************************************************************************
'概　要：対象求人情報の全メールテンプレートの編集ページへのリンクを取得
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_GetMyOrderで生成されたレコードセットオブジェクト
'　　　：vOrderCode		：
'　　　：vMTCnt			：現在のメールテンプレート件数
'作成者：Lis Kokubo
'作成日：2007/06/15
'備　考：
'使用元：しごとナビ/mailtemplate/manager.asp
'******************************************************************************
Function GetMailTemplateLink(ByRef rDB, ByVal vUserCode, ByVal vOrderCode, ByVal vMTCnt)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim iSEQ
	Dim sMTTName
	Dim sSubject

	Dim sParam
	Dim sEditLink
	Dim sEditButton
	Dim sDelButton

	If vMTCnt < 5 Then
		GetMailTemplateLink = "<input type=""button"" value=""新規作成"" onclick=""location.href = '" & HTTPS_NAVI_CURRENTURL & "mailtemplate/regist.asp?ordercode=" & vOrderCode & "';""><br>"
	Else
		GetMailTemplateLink = "<p class=""m0"" style=""font-size:10px; color:#ff0000;"">テンプレートの数が上限に達しています。</p>"
	End If

	sSQL = "up_GetListMailTemplate '" & vUserCode & "', '" & vOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		iSEQ = oRS.Collect("SEQ")
		sMTTName = ChkStr(oRS.Collect("MailTemplateTypeName"))
		sSubject = ChkStr(oRS.Collect("Subject"))
		If Len(sSubject) > 20 Then sSubject = Left(sSubject, 20) & "..."

		sParam = "?ordercode=" & vOrderCode & "&amp;seq=" & iSEQ

		sEditButton = "<input type=""button"" value=""ｺﾋﾟｰ"" style=""width:35px;"" onclick=""location.href='" & HTTPS_NAVI_CURRENTURL & "mailtemplate/copy.asp" & sParam & "';"">"
		sDelButton = "<input type=""button"" value=""削除"" style=""width:35px;"" onclick=""location.href='" & HTTPS_NAVI_CURRENTURL & "mailtemplate/delete.asp" & sParam & "';"">"

		sEditLink = "<div style=""float:left; width:100px;"">" & sMTTName & "</div>"
		sEditLink = sEditLink & "<div style=""float:left; width:230px;""><a href=""" & HTTPS_NAVI_CURRENTURL & "mailtemplate/regist.asp" & sParam & """>" & sSubject & "</a></div>"
		sEditLink = sEditLink & "<div style=""float:left; width:80px;"">" & sEditButton & "&nbsp;" & sDelButton & "</div>"
		sEditLink = sEditLink & "<div style=""clear:both;""></div>"

		GetMailTemplateLink = GetMailTemplateLink & "<div style="" border-bottom:1px dashed #ccc; margin:2px 0; padding:2px 0;"">" & sEditLink & "</div>"
		oRS.MoveNext
	Loop
	Call RSClose(oRS)
End Function

'******************************************************************************
'概　要：メールテンプレートの入力画面出力
'引　数：vModeText				："新規作成" or "編集"
'　　　：vOrderCode				：メールテンプレート作成対象の情報コード
'　　　：vSEQ					：メールテンプレート作成対象の情報コードの連番
'　　　：vMailTemplateTypeCode	：メールテンプレート種類コード
'　　　：vSubject				：件名
'　　　：vBody					：内容
'　　　：rErrStyle				：ディクショナリ：入力エラー時のスタイルシートを保持したもの
'作成者：Lis Kokubo
'作成日：2007/06/15
'備　考：
'使用元：しごとナビ/mailtemplate/regist.asp
'******************************************************************************
Function DspInputMailTemplate(ByVal vModeText, ByVal vOrderCode, ByVal vSEQ, ByVal vMailTemplateTypeCode, ByVal vSubject, ByVal vBody, ByRef rErrStyle)
%>
	<form id="frmmailtemplate" action="<%= HTTPS_NAVI_CURRENTURL %>mailtemplate/regist.asp?ordercode=<%= vOrderCode %>&amp;seq=<%= vSEQ %>" method="post">
	<input id="regmode" name="frmregmode" type="hidden" value="1">

	<p>
		<input type="button" value="他メールテンプレートの内容をコピー" style="width:200px;" onclick="window.open('<%= HTTPS_NAVI_CURRENTURL %>mailtemplate/copyreference.asp', 'mywindow6', 'width=635, height=500, menubar=no, toolbar=no, scrollbars=yes');">
		<span style="font-size:10px;">・・・他の求人票のメールテンプレートの内容をコピーして反映させる</span>
	</p>

	<table class="pattern8" border="0">
		<thead>
			<tr>
				<th colspan="2">メールテンプレート<%= vModeText %></th>
			</tr>
		</thead>
		<tbody>
			<tr>
				<th style="width:138px; text-align:center;">情報コード</th>
				<td style="width:439px;"><%= vOrderCode %></td>
			</tr>
			<tr>
				<th class="need" style="text-align:center;">テンプレート種類</th>
				<td style="<%= rErrStyle("frmmailtemplatetypecode") %>">
					<%= GetRadioHtml("MailTemplateType", vMailTemplateTypeCode, "frmmailtemplatetypecode") %>
				</td>
			</tr>
			<tr>
				<th class="need" style="text-align:center;">
					メール件名<br>
					<a href="<%= HTTP_NAVI_CURRENTURL %>counter.htm" target="_brank" style="font-size:10px;">文字数カウント</a>
				</th>
				<td><input id="subject" name="frmsubject" type="text" value="<%= vSubject %>" style="width:100%;<%= rErrStyle("frmsubject") %>"><br><p class="m0">※全角５０文字以内</p></td>
			</tr>
			<tr>
				<th class="need" style="text-align:center;">
					メール内容<br>
					<a href="<%= HTTP_NAVI_CURRENTURL %>counter.htm" target="_brank" style="font-size:10px;">文字数カウント</a>
				</th>
				<td><textarea id="body" name="frmbody" style="width:100%; height:400px;<%= rErrStyle("frmbody") %>"><%= vBody %></textarea><br><p class="m0">※全角２０００文字以内</p></td>
			</tr>
		</tbody>
	</table>
	<br>
<%
	If vSEQ <> "" Then
		'編集画面の場合は、登録ボタン、コピーボタン、削除ボタンを配置する。
%>
	<div style="width:600px;">
		<div align="center" style="float:left; width:150px; border-bottom:1px dotted #666666;"><input type="button" value="確　認" style="width:80px;" onclick="document.forms.frmmailtemplate.submit();"></div>
		<div align="left" style="float:left; width:450px; border-bottom:1px dotted #666666;">→編集した内容を確認し、登録します。</div>
		<div style="clear:both;"></div>
	</div>
	<div style="width:600px; padding-top:10px;">
		<div align="center" style="float:left; width:150px; border-bottom:1px dotted #666666;"><input type="button" value="コピー" style="width:80px;" onclick="location.href='<%= HTTPS_NAVI_CURRENTURL %>mailtemplate/copy.asp?ordercode=<%= qsOrderCode %>&amp;seq=<%= qsSEQ %>';"></div>
		<div align="left" style="float:left; width:450px; border-bottom:1px dotted #666666;">→このテンプレート（<span style="color:#ff0000;">編集前</span>）を、他の求人票にコピーします。</div>
		<div style="clear:both;"></div>
	</div>
	<div style="width:600px;">
		<div align="center" style="float:left; width:150px; border-bottom:1px dotted #666666;"><input type="button" value="削　除" style="width:80px;" onclick="location.href='<%= HTTPS_NAVI_CURRENTURL %>mailtemplate/delete.asp?ordercode=<%= qsOrderCode %>&amp;seq=<%= qsSEQ %>';"></div>
		<div align="left" style="float:left; width:450px; border-bottom:1px dotted #666666;">→このテンプレートを削除します。</div>
		<div style="clear:both;"></div>
	</div>
	<br>
<%
	Else
		'新規作成の場合は、登録ボタンのみを配置する。
%>
	<div align="center"><input type="submit" value="確　認"></div>
<%
	End If
%>
	</form>
<%
End Function

'******************************************************************************
'概　要：メールテンプレートの入力内容確認画面出力
'引　数：vModeText				："新規作成" or "編集"
'　　　：vOrderCode				：メールテンプレート作成対象の情報コード
'　　　：vSEQ					：メールテンプレート作成対象の情報コードの連番
'　　　：vMailTemplateTypeCode	：メールテンプレート種類コード
'　　　：vSubject				：件名
'　　　：vBody					：内容
'　　　：rErrStyle				：ディクショナリ：入力エラー時のスタイルシートを保持したもの
'作成者：Lis Kokubo
'作成日：2007/06/15
'備　考：
'使用元：しごとナビ/mailtemplate/regist.asp
'******************************************************************************
Function DspConfMailTemplate(ByVal vModeText, ByVal vOrderCode, ByVal vSEQ, ByVal vMailTemplateTypeCode, ByVal vSubject, ByVal vBody)
	Session("regmailtemplate") = "1"
%>
	<form id="frmmailtemplate" action="<%= HTTPS_NAVI_CURRENTURL %>mailtemplate/regist.asp?ordercode=<%= vOrderCode %>&amp;seq=<%= vSEQ %>" method="post">
	<input id="regmode" name="frmregmode" type="hidden" value="2">
	<input id="mailtemplatetypecode" name="frmmailtemplatetypecode" type="hidden" value="<%= vMailTemplateTypeCode %>">
	<input id="subject" name="frmsubject" type="hidden" value="<%= vSubject %>">
	<input id="body" name="frmbody" type="hidden" value="<%= vBody %>">
	<table class="pattern8" border="0">
		<thead>
			<tr>
				<th colspan="2">メールテンプレート<%= vModeText %></th>
			</tr>
		</thead>
		<tbody>
			<tr>
				<th style="width:138px; text-align:center;">情報コード</th>
				<td style="width:439px;"><%= vOrderCode %></td>
			</tr>
			<tr>
				<th class="need" style="text-align:center;">テンプレート種類</th>
				<td><%= GetDetail("MailTemplateType", vMailTemplateTypeCode) %></td>
			</tr>
			<tr>
				<th class="need" style="text-align:center;">メール件名</th>
				<td><%= ChgSQLtoView(vSubject) %></td>
			</tr>
			<tr>
				<th class="need" style="text-align:center;">
					メール内容
				</th>
				<td><%= ChgSQLtoView(vBody) %></td>
			</tr>
		</tbody>
	</table>
	<br>
	<div align="center"><input type="submit" value="登　録"></div>
	</form>
<%
End Function

'******************************************************************************
'概　要：メールテンプレート入力画面のコピー元選択部分のテンプレート一覧を出力
'引　数：rDB				：接続中のDBConnection
'　　　：vUserCode			：ログイン中ユーザ
'　　　：vContactPersonName	：求人担当者フィルタ
'　　　：vPageSize			：１ページあたりの表示件数
'　　　：vPage				：表示中ページ
'作成者：Lis Kokubo
'作成日：2007/06/15
'備　考：
'使用元：しごとナビ/mailtemplate/regist.asp
'******************************************************************************
Function DspMailTemplateRefList(ByRef rDB, ByVal vUserCode, ByVal vContactPersonName, ByVal vPageSize, ByVal vPage)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderCode
	Dim sHtmlPageCtrl
	Dim iRow
	Dim sFilterPerson	'求人担当者一覧（<option></option> を保持する）

	sSQL = "up_GetListMailTemplateExists '" & vUserCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	'求人担当者一覧を取得
	If GetRSState(oRS) = True Then
		sFilterPerson = GetContactPersonNameMTOptionHtml(rDB, G_USERID, vContactPersonName)
	End If

	'求人担当者で絞込み
	If GetRSState(oRS) = True And vContactPersonName <> "" Then
		If vContactPersonName <> "" Then oRS.Filter = "ContactPersonName = '" & vContactPersonName & "'"
	End If

	'ページコントロール取得
	If GetRSState(oRS) = True Then
		sHtmlPageCtrl = GetPageControlHtml(rDB, oRS, vPageSize, vPage)
	End If

	Response.Write sHtmlPageCtrl
%>
	<table class="pattern8" border="0">
		<thead>
			<tr>
				<th colspan="3">参照するメールテンプレートを選択してください</th>
			</tr>
		</thead>
		<tbody>
			<tr>
				<th style="width:68px; text-align:center;">情報コード</th>
				<th style="width:109px;">
					<select id="contactpersonname" name="frmcontactpersonname" style="width:109px;" onchange="ChgPage(1);">
						<option value="">(求人担当者)</otpion>
						<%= sFilterPerson %>
					</select>
				</th>
				<th style="width:389px;">メールテンプレート種類</th>
			</tr>
<%
	iRow = 1
	If GetRSState(oRS) = True Then
		Do While GetRSState(oRS) = True And iRow <= vPageSize
			Call DspMailTemplateRefListOne(rDB, oRS, G_USERID)
			oRS.MoveNext
			iRow = iRow + 1
		Loop
	Else
%>
			<tr>
				<td colspan="3">求人票がありません。</td>
			</tr>
<%
	End If
%>
		</tbody>
	</table>
<%
	Response.Write sHtmlPageCtrl
	Call RSClose(oRS)
End Function

'******************************************************************************
'概　要：メールテンプレート一覧の個々のテンプレートを出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_GetMyOrderで生成されたレコードセットオブジェクト
'作成者：Lis Kokubo
'作成日：2007/06/15
'備　考：
'使用元：しごとナビ/mailtemplate/manager.asp
'******************************************************************************
Function DspMailTemplateRefListOne(ByRef rDB, ByRef rRS, ByVal vUserCode)
	Dim sOrderCode
	Dim sContactPersonName
	Dim sAnc				'求人票へのリンク

	If GetRSState(rRS) = False Then Exit Function

	sOrderCode = ChkStr(rRS.Collect("OrderCode"))
	sContactPersonName = ChkStr(rRS.Collect("ContactPersonName"))

	sAnc = "<a href=""" & HTTP_NAVI_CURRENTURL & "order/order_detail.asp?ordercode=" & sOrderCode & """ target=""_brank"">" & sOrderCode & "</a>"
%>
			<tr>
				<td style="text-align:center;"><%= sAnc %></td>
				<td><%= sContactPersonName %></td>
				<td>
					<%= GetMailTemplateRefLink(rDB, vUserCode, sOrderCode) %>
				</td>
			</tr>
<%
End Function

'******************************************************************************
'概　要：対象求人情報の全メールテンプレートの編集ページへのリンクを取得
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_GetMyOrderで生成されたレコードセットオブジェクト
'作成者：Lis Kokubo
'作成日：2007/06/15
'備　考：
'使用元：しごとナビ/mailtemplate/manager.asp
'******************************************************************************
Function GetMailTemplateRefLink(ByRef rDB, ByVal vUserCode, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim iSEQ
	Dim sMTTName
	Dim sSubject

	Dim sParam
	Dim sEditLink

	GetMailTemplateRefLink = ""

	sSQL = "up_GetListMailTemplate '" & vUserCode & "', '" & vOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		iSEQ = oRS.Collect("SEQ")
		sMTTName = ChkStr(oRS.Collect("MailTemplateTypeName"))
		sSubject = ChkStr(oRS.Collect("Subject"))
		If Len(sSubject) > 25 Then sSubject = Left(sSubject, 25) & "..."

		sParam = "?ordercode=" & vOrderCode & "&amp;seq=" & iSEQ & "&amp;conf=1"
		sEditLink = "<div style=""float:left; width:80px;"">" & sMTTName & "</div>" & _
			"<div style=""float:left; width:309px;"">" & "<a href=""" & HTTPS_NAVI_CURRENTURL & "mailtemplate/copyreference.asp" & sParam & """>" & sSubject & "</a></div>" & _
			"<div style=""clear:both;""></div>"

		GetMailTemplateRefLink = GetMailTemplateRefLink & _
			"<div>" & sEditLink & "</div>"
		oRS.MoveNext
	Loop
	Call RSClose(oRS)
End Function

'******************************************************************************
'概　要：メールテンプレートの入力内容確認画面出力
'引　数：rDB		：接続中ＤＢオブジェクト
'　　　：vUserCode	：ログイン中ユーザコード
'　　　：vOrderCode	：メールテンプレート作成対象の情報コード
'　　　：vSEQ		：メールテンプレート作成対象の情報コードの連番
'作成者：Lis Kokubo
'作成日：2007/06/15
'備　考：
'使用元：しごとナビ/mailtemplate/regist.asp
'******************************************************************************
Function DspMailTemplateRefCopy(ByRef rDB, ByVal vUserCode, ByVal vOrderCode, ByVal vSEQ)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sMailTemplateTypeCode
	Dim sSubject
	Dim sBody

	DspMailTemplateRefCopy = False

	sSQL = "up_GetDetailMailTemplate '" & vUserCode & "', '" & vOrderCode & "', '" & vSEQ & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	If GetRSState(oRS) = False Then Exit Function

	sMailTemplateTypeCode = ChkStr(oRS.Collect("MailTemplateTypeCode"))
	sSubject = ChkStr(oRS.Collect("Subject"))
	sBody = ChkStr(oRS.Collect("Body"))
%>
	<input id="mailtemplatetypecode" name="frmmailtemplatetypecode" type="hidden" value="<%= sMailTemplateTypeCode %>">
	<input id="subject" name="frmsubject" type="hidden" value="<%= sSubject %>">
	<input id="body" name="frmbody" type="hidden" value="<%= sBody %>">
	<table class="pattern8" border="0">
		<thead>
			<tr>
				<th colspan="2">コピーするメールテンプレート</th>
			</tr>
		</thead>
		<tbody>
			<tr>
				<th style="width:138px; text-align:center;">情報コード</th>
				<td style="width:439px;"><%= vOrderCode %></td>
			</tr>
			<tr>
				<th style="text-align:center;">テンプレート種類</th>
				<td><%= GetDetail("MailTemplateType", sMailTemplateTypeCode) %></td>
			</tr>
			<tr>
				<th style="text-align:center;">メール件名</th>
				<td><%= ChgSQLtoView(sSubject) %></td>
			</tr>
			<tr>
				<th style="text-align:center;">
					メール内容
				</th>
				<td><%= ChgSQLtoView(sBody) %></td>
			</tr>
		</tbody>
	</table>
	<br>
	<div align="center"><input type="button" value="コピー" onclick="setMailTemplateCopy();"></div>

<script type="text/javascript" language="javascript">
//<!--
function setMailTemplateCopy(){
	var oopfrm = opener.document.forms.frmmailtemplate;
	var oopmttcode = opener.document.getElementsByName('frmmailtemplatetypecode');

	//get data
	var smttcode = document.getElementById('mailtemplatetypecode').value
	var ssubject = document.getElementById('subject').value
	var sbody = document.getElementById('body').value

	//set openr
	for(var idx = 0; oopmttcode[idx] != null; idx++){
		if(oopmttcode[idx].value == smttcode){
			oopmttcode[idx].checked = true;
			break;
		}
	}
	oopfrm.subject.value = ssubject;
	oopfrm.body.value = sbody;
	close();
}
//-->
</script>
<%
	DspMailTemplateRefCopy = True
End Function

'******************************************************************************
'概　要：メールテンプレートを保持する求人票の担当者一覧を <option></option> 形式で取得
'引　数：rDB		：接続中のＤＢオブジェクト
'　　　：vUserCode	：ログイン中ユーザコード
'　　　：vPersonName：絞り込む求人票の担当者名
'作成者：Lis Kokubo
'作成日：2007/06/15
'備　考：
'使用元：しごとナビ/mailtemplate/regist.asp
'******************************************************************************
Function GetContactPersonNameMTOptionHtml(ByRef rDB, ByVal vUserCode, ByVal vPersonName)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	GetContactPersonNameMTOptionHtml = ""

	sSQL = "up_GetListContactPersonNameMT '" & vUserCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		If oRS.Collect("PersonName") = vPersonName Then
			GetContactPersonNameMTOptionHtml = GetContactPersonNameMTOptionHtml & _
				"<option value=""" & oRS.Collect("PersonName") & """ selected>" & oRS.Collect("PersonName") & "</option>"
		Else
			GetContactPersonNameMTOptionHtml = GetContactPersonNameMTOptionHtml & _
				"<option value=""" & oRS.Collect("PersonName") & """>" & oRS.Collect("PersonName") & "</option>"
		End If
		oRS.MoveNext
	Loop
End Function

'******************************************************************************
'概　要：メールテンプレートＤＢ登録～登録完了・失敗画面出力
'引　数：vModeText				："新規作成" or "編集"
'　　　：vUserCode				：ログイン中ユーザコード
'　　　：vOrderCode				：メールテンプレート作成対象の情報コード
'　　　：vSEQ					：メールテンプレート作成対象の情報コードの連番
'　　　：vMailTemplateTypeCode	：メールテンプレート種類コード
'　　　：vSubject				：件名
'　　　：vBody					：内容
'作成者：Lis Kokubo
'作成日：2007/06/15
'備　考：
'使用元：しごとナビ/mailtemplate/regist.asp
'******************************************************************************
Function DspRegMailTemplate(ByRef rDB, ByVal vModeText, ByVal vUserCode, ByVal vOrderCode, ByVal vSEQ, ByVal vMailTemplateTypeCode, ByVal vSubject, ByVal vBody)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim flgReg

	flgReg = False
	If Session("regmailtemplate") = "1" Then
		'登録処理
		flgReg = RegMailTemplate(rDB, vUserCode, vOrderCode, vSEQ, vMailTemplateTypeCode, vSubject, vBody)
	End If
	Session.Contents.Remove("regmailtemplate")

	If flgReg = True Then
		Response.Write "<p><b>メールテンプレートを登録しました。</b></p>"
	Else
		Response.Write "<p><b>メールテンプレートの登録に<span style=""color:#ff0000;"">失敗</span>しました。</b></p>"
	End If

	Response.Write "<p><a href=""" & HTTP_NAVI_CURRENTURL & "order/order_detail.asp?ordercode=" & vOrderCode & """>求人票詳細ページへ</a></p>"
	Response.Write "<p><a href=""" & HTTP_NAVI_CURRENTURL & "mailtemplate/manager.asp"">メールテンプレート管理ページへ</a></p>"
End Function

'******************************************************************************
'概　要：メールテンプレートの入力内容確認画面出力
'引　数：rDB		：接続中ＤＢオブジェクト
'　　　：vUserCode	：ログイン中ユーザコード
'　　　：vOrderCode	：メールテンプレート作成対象の情報コード
'　　　：vSEQ		：メールテンプレート作成対象の情報コードの連番
'作成者：Lis Kokubo
'作成日：2007/06/15
'備　考：
'使用元：しごとナビ/mailtemplate/regist.asp
'******************************************************************************
Function DspConfDeleteMailTemplate(ByRef rDB, ByVal vUserCode, ByVal vOrderCode, ByVal vSEQ)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sMailTemplateTypeCode
	Dim sSubject
	Dim sBody

	DspConfDeleteMailTemplate = False

	sSQL = "up_GetDetailMailTemplate '" & vUserCode & "', '" & vOrderCode & "', '" & vSEQ & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	If GetRSState(oRS) = False Then Exit Function

	sMailTemplateTypeCode = ChkStr(oRS.Collect("MailTemplateTypeCode"))
	sSubject = ChkStr(oRS.Collect("Subject"))
	sBody = ChkStr(oRS.Collect("Body"))
%>
	<form id="frmmailtemplate" action="<%= HTTPS_NAVI_CURRENTURL %>mailtemplate/delete.asp?ordercode=<%= vOrderCode %>&amp;seq=<%= vSEQ %>" method="post">
	<input id="regmode" name="frmregmode" type="hidden" value="1">
	<table class="pattern8" border="0">
		<thead>
			<tr>
				<th colspan="2">メールテンプレート削除確認</th>
			</tr>
		</thead>
		<tbody>
			<tr>
				<th style="width:138px; text-align:center;">情報コード</th>
				<td style="width:439px;"><%= vOrderCode %></td>
			</tr>
			<tr>
				<th style="text-align:center;">テンプレート種類</th>
				<td><%= GetDetail("MailTemplateType", sMailTemplateTypeCode) %></td>
			</tr>
			<tr>
				<th style="text-align:center;">メール件名</th>
				<td><%= sSubject %></td>
			</tr>
			<tr>
				<th style="text-align:center;">
					メール内容<br>
				</th>
				<td><%= Replace(sBody, vbCrLf, "<br>") %></td>
			</tr>
		</tbody>
	</table>
	<br>
	<div align="center"><input type="submit" value="削　除"></div>
	</form>
<%
	DspConfDeleteMailTemplate = True
End Function

'******************************************************************************
'概　要：メールテンプレートのコピー元を出力
'引　数：rDB				：接続中のDBConnection
'　　　：vUserCode			：ログイン中ユーザ
'　　　：vContactPersonName	：求人担当者フィルタ
'　　　：vPageSize			：１ページあたりの表示件数
'　　　：vPage				：表示中ページ
'作成者：Lis Kokubo
'作成日：2007/06/18
'備　考：
'使用元：しごとナビ/mailtemplate/manager.asp
'******************************************************************************
Function DspCopyMailTemplate(ByRef rDB, ByVal vUserCode, ByVal vOrderCode, ByVal vSEQ)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sMailTemplateTypeCode
	Dim sSubject
	Dim sBody

	DspCopyMailTemplate = False

	sSQL = "up_GetDetailMailTemplate '" & vUserCode & "', '" & vOrderCode & "', '" & vSEQ & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	If GetRSState(oRS) = False Then Exit Function

	sMailTemplateTypeCode = ChkStr(oRS.Collect("MailTemplateTypeCode"))
	sSubject = ChkStr(oRS.Collect("Subject"))
	sBody = ChkStr(oRS.Collect("Body"))
%>
	<table class="pattern8" border="0">
		<thead>
			<tr>
				<th colspan="2">コピー元メールテンプレート</th>
			</tr>
		</thead>
		<tbody>
			<tr>
				<th style="width:138px; text-align:center;">情報コード</th>
				<td style="width:439px;"><%= vOrderCode %></td>
			</tr>
			<tr>
				<th style="text-align:center;">テンプレート種類</th>
				<td><%= GetDetail("MailTemplateType", sMailTemplateTypeCode) %></td>
			</tr>
			<tr>
				<th style="text-align:center;">メール件名</th>
				<td><%= sSubject %></td>
			</tr>
			<tr>
				<th style="text-align:center;">
					メール内容<br>
				</th>
				<td><%= Replace(sBody, vbCrLf, "<br>") %></td>
			</tr>
		</tbody>
	</table>
	<br>
<%
	DspCopyMailTemplate = True
End Function

'******************************************************************************
'概　要：メールテンプレートコピー画面の一覧部分出力
'引　数：rDB				：接続中のDBConnection
'　　　：vUserCode			：ログイン中ユーザ
'　　　：vContactPersonName	：求人担当者フィルタ
'　　　：vPageSize			：１ページあたりの表示件数
'　　　：vPage				：表示中ページ
'作成者：Lis Kokubo
'作成日：2007/06/15
'備　考：
'使用元：しごとナビ/mailtemplate/manager.asp
'******************************************************************************
Function DspCopyMailTemplateList(ByRef rDB, ByVal vUserCode, ByVal vOrderCode, ByVal vSEQ, ByVal vPageSize, ByVal vPage, ByVal vContactPersonName)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dicOrderDetail	'ディクショナリ：求人票詳細
	Dim sOrderCode		'情報コード
	Dim sFilterPerson	'求人票一覧のレコードセットを求人担当者で絞り込むフィルター
	Dim sHtmlPageCtrl	'ページコントロールのHTML
	Dim iRow

	sSQL = "up_GetMyOrder '" & vUserCode & "', '0', '', ''"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	'求人担当者一覧を取得
	If GetRSState(oRS) = True Then
		sFilterPerson = GetContactPersonNameMTOptionHtml(rDB, G_USERID, vContactPersonName)
	End If

	'求人担当者で絞込み
	If GetRSState(oRS) = True And vContactPersonName <> "" Then
		If vContactPersonName <> "" Then oRS.Filter = "ContactPersonName = '" & vContactPersonName & "'"
	End If

	'ページコントロール取得
	If GetRSState(oRS) = True Then
		sHtmlPageCtrl = GetPageControlHtml(rDB, oRS, vPageSize, vPage)
		sHtmlPageCtrl = "<div style=""border-top:1px dotted #666666; border-bottom:1px dotted #666666;"">" & sHtmlPageCtrl & "</div>"
	End If
%>
<select id="contactpersonname" name="frmcontactpersonname" style="margin-bottom:5px;" onchange="this.form.frmpage.value='1'; ChgPage(1);">
	<option value="">----- コピー先求人票の担当者 -----</otpion>
	<%= sFilterPerson %>
</select>
<%

	If GetRSState(oRS) = True Then
		iRow = 1
		Response.Write sHtmlPageCtrl

		Do While GetRSState(oRS) = True And iRow <= 10
			sOrderCode = oRS.Collect("OrderCode")
			Set dicOrderDetail = GetDicOrderDetail(rDB, sOrderCode)
			Call DspMailTemplateListOne2(dicOrderDetail, vOrderCode, vSEQ)
			oRS.MoveNext
			iRow = iRow + 1
			Set dicOrderDetail = Nothing
		Loop

		Response.Write sHtmlPageCtrl
	Else
		Response.Write "<p>求人票がありません。</p>"
	End If

	Call RSClose(oRS)
End Function

'******************************************************************************
'概　要：メールテンプレートコピー画面のコピー先求人票一覧を出力
'引　数：rDic		：
'　　　：vOrderCode	：
'　　　：vSEQ		：
'作成者：Lis Kokubo
'作成日：2007/06/18
'備　考：
'使用元：しごとナビ/mailtemplate/manager.asp
'******************************************************************************
Function DspMailTemplateListOne2(ByRef rDic, ByVal vOrderCode, ByVal vSEQ)
	Dim sSalary

	DspMailTemplateListOne2 = False

	If IsObject(rDic) = False Then Exit Function
	If Len(rDic("OrderCode")) = 0 Then Exit Function

	sSalary = ""
%>
<table class="cw" border="0" style="margin:2px 0px;">
	<tbody>
	<tr>
		<td style="width:125px; padding-right:5px; vertical-align:top;">
			→&nbsp;
<%
	If rDic("MailTemplateCnt") < 5 Then
%>
			<a href="<%= HTTPS_NAVI_CURRENTURL %>mailtemplate/copy.asp?ordercode=<%= vOrderCode %>&amp;seq=<%= vSEQ %>&amp;copyto=<%= rDic("OrderCode") %>">この求人票へコピー</a>
<%
	Else
%>
			<span style="font-size:10px; color:#ff0000;">制限数に達しています</span>
<%
	End If
%>
		</td>
		<td style="width:480px; vertical-align:top;"><%= rDic("OrderCode") & "&nbsp;：&nbsp;" & rDic("JobTypeDetail") & rDic("WorkingType") %></td>
	</tr>
	</tbody>
</table>
<%
	DspMailTemplateListOne2 = True
End Function

'******************************************************************************
'概　要：メールテンプレートコピーＤＢ登録～登録完了画面出力
'引　数：rDB		：接続中ＤＢオブジェクト
'　　　：vUserCode	：ログイン中ユーザコード
'　　　：vOrderCode	：メールテンプレート作成対象の情報コード
'　　　：vSEQ		：メールテンプレート作成対象の情報コードの連番
'　　　：vCopyTo	：コピー先情報コード
'作成者：Lis Kokubo
'作成日：2007/06/15
'備　考：
'使用元：しごとナビ/mailtemplate/regist.asp
'******************************************************************************
Function DspRegCopyMailTemplate(ByRef rDB, ByVal vUserCode, ByVal vOrderCode, ByVal vSEQ, ByVal vCopyTo)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim flgReg

	sSQL = "up_GetDetailMailTemplate '" & G_USERID & "', '" & qsOrderCode & "', '" & qsSEQ & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		sMailTemplateTypeCode = oRS.Collect("MailTemplateTypeCode")
		sSubject = oRS.Collect("Subject")
		sBody = oRS.Collect("Body")
	End If

	flgReg = False
	If Session("regmailtemplate") = "1" Then
		'登録処理
		flgReg = RegMailTemplate(rDB, vUserCode, vCopyTo, "", sMailTemplateTypeCode, sSubject, sBody)
	End If
	Session.Contents.Remove("regmailtemplate")

	If flgReg = True Then
		Response.Write "<p><b>メールテンプレートをコピーしました。</b></p>"
	Else
		Response.Write "<p><b>メールテンプレートのコピーに<span style=""color:#ff0000;"">失敗</span>しました。</b></p>"
	End If

	Response.Write "<p><a href=""" & HTTP_NAVI_CURRENTURL & "order/order_detail.asp?ordercode=" & qsOrderCode & """>求人票詳細ページへ</a></p>"
	Response.Write "<p><a href=""" & HTTP_NAVI_CURRENTURL & "mailtemplate/manager.asp"">メールテンプレート管理ページへ</a></p>"
End Function

'******************************************************************************
'概　要：メールテンプレートＤＢ削除～削除完了画面出力
'引　数：rDB		：接続中ＤＢオブジェクト
'　　　：vUserCode	：ログイン中ユーザコード
'　　　：vOrderCode	：メールテンプレート作成対象の情報コード
'　　　：vSEQ		：メールテンプレート作成対象の情報コードの連番
'作成者：Lis Kokubo
'作成日：2007/06/15
'備　考：
'使用元：しごとナビ/mailtemplate/regist.asp
'******************************************************************************
Function DspDeleteMailTemplate(ByRef rDB, ByVal vUserCode, ByVal vOrderCode, ByVal vSEQ)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim flgDel

	flgDel = DelMailTemplate(rDB, vUserCode, vOrderCode, vSEQ)

	If flgDel = True Then
		Response.Write "<p><b>メールテンプレートを削除しました。</b></p>"
	Else
		Response.Write "<p><b>メールテンプレートの削除に<span style=""color:#ff0000;"">失敗</span>しました。</b></p>"
	End If

	Response.Write "<p><a href=""" & HTTP_NAVI_CURRENTURL & "order/order_detail.asp?ordercode=" & vOrderCode & """>求人票詳細ページへ</a></p>"
	Response.Write "<p><a href=""" & HTTP_NAVI_CURRENTURL & "mailtemplate/manager.asp"">メールテンプレート管理ページへ</a></p>"

	DspDeleteMailTemplate = True
End Function

'******************************************************************************
'概　要：メールテンプレートの登録処理
'引　数：rDB					：接続中のDBConnection
'　　　：vUserCode				：
'　　　：vOrderCode				：
'　　　：vSEQ					：
'　　　：vMailTemplateTypeCode	：
'　　　：vSubject				：
'　　　：sBody					：
'作成者：Lis Kokubo
'作成日：2007/06/15
'備　考：
'使用元：しごとナビ/mailtemplate/regist.asp
'******************************************************************************
Function RegMailTemplate(ByRef rDB, ByVal vUserCode, ByVal vOrderCode, ByVal vSEQ, ByVal vMailTemplateTypeCode, ByVal vSubject, ByVal vBody)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim flgData

	RegMailTemplate = False

	sSQL = "up_Reg_C_MailTemplate '" & vOrderCode & "', '" & vSEQ & "', '" & vMailTemplateTypeCode & "', '" & ChkSQLStr(vSubject) & "', '" & ChkSQLStr(vBody) & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	If flgQE = True Then RegMailTemplate = True
End Function

'******************************************************************************
'概　要：メールテンプレートの削除処理
'引　数：rDB					：接続中のDBConnection
'　　　：vUserCode				：
'　　　：vOrderCode				：
'　　　：vSEQ					：
'　　　：vMailTemplateTypeCode	：
'　　　：vSubject				：
'　　　：sBody					：
'作成者：Lis Kokubo
'作成日：2007/06/15
'備　考：
'使用元：しごとナビ/mailtemplate/regist.asp
'******************************************************************************
Function DelMailTemplate(ByRef rDB, ByVal vUserCode, ByVal vOrderCode, ByVal vSEQ)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim flgData

	DelMailTemplate = False

	sSQL = "up_Del_C_MailTemplate '" & vUserCode & "', '" & vOrderCode & "', '" & vSEQ & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	If flgQE = True Then DelMailTemplate = True
End Function
%>
