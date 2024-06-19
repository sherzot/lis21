<%
'**********************************************************************************************************************
'概　要：メール詳細ページ /staff/mailhistory_person_entity.asp
'　　　：上記ページで出力用の関数群をこのファイルに用意する。
'　　　：
'　　　：■■■　前提条件　■■■
'　　　：要事前インクルード
'　　　：/config/personel.asp
'　　　：/include/commonfunc.asp
'一　覧：■■■　メール詳細ページ出力用　■■■
'　　　：DspMailReturnBtn	：返信ボタン出力
'　　　：DspMailDetail		：メール詳細を出力
'　　　：DspNoMailDetail	：メールが無い場合の文言出力
'**********************************************************************************************************************

'******************************************************************************
'概　要：返信ボタンを出力
'引　数：vMode		：現在の表示モード ["0"]受信メール一覧 ["1"]送信メール一覧
'　　　：vAnswerNG	：返信可否 ["0"]返信可 ["1"]返信不可
'作成者：Lis Kokubo
'作成日：2007/03/02
'備　考：
'使用元：ナビ/staff/maildetail_person_entity.asp
'******************************************************************************
Function DspMailReturnBtnCustom(ByVal vMode, ByVal vAnswerNG)
	If vMode <> "1" Then
		If vAnswerNG = "0" Then
			Response.Write "<br><div align=""center""><input type=""button"" value=""返　信"" onclick=""SendAnswer();""></div><br>"
		Else
			Response.Write "<b>この求人票は掲載を終了しており、連絡を取ることができません。</b><br><br>"
		End If
	End If
End Function

'******************************************************************************
'概　要：メール詳細を出力
'引　数：vMode		：現在の表示モード ["0"]受信メール一覧 ["1"]送信メール一覧
'　　　：vAnswerNG	：返信可否 ["0"]返信可 ["1"]返信不可
'作成者：Lis Kokubo
'作成日：2007/03/02
'備　考：
'使用元：ナビ/staff/maildetail_person_entity.asp
'******************************************************************************
Function DspMailDetailCustom(ByRef rRS, ByVal vMode, ByVal vAnswerNG)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
%>
<table class="pattern1" border="0" style="width:600px;">
	<thead>
	<tr>
		<th colspan="2" style="width:588px;"><%= rRS.Collect("Subject") %></th>
	</tr>
	</thead>
	<tbody>
	<tr>
<%
	If vMode = "1" Then
		'送信画面の場合は宛先
%>
		<th style="width:188px;">宛先</th>
<%
	Else
		'開封済みにする
		sSQL = "sp_Reg_MailOpenDay '" & rRS.Collect("ID") & "'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		'受信画面の場合は差出人
%>
		<th style="width:188px;">差出人</th>
<%
	End If
%>
		<td style="width:389px;">
<%
	Response.Write rRS.Collect("CompanyName")
	If sAnswerNG = "0" And Trim(rRS.Collect("OrderCode")) <> "" Then
		Response.Write "　　<a href=""../order/order_detail.asp?OrderCode=" & rRS.Collect("OrderCode") & """>求人票詳細</a>"
	ElseIf sAnswerNG = "1" Then
		Response.Write "　　" & rRS.Collect("OrderCode")
	End If
%>
		</td>
	</tr>
	<tr>
		<th>内容</th>
		<td>
			<textarea rows="20" cols="66" readonly style="height:300px;"><%= rRS.Collect("Body") %></textarea>
		</td>
	</tr>
	</tbody>
</table>
<br>
<%
End Function

Function DspNoMailDetail()
	Response.Write "<b>指定されたメールは存在しないか、削除されています。</b><br><br>"
End Function
%>
