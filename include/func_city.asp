<%
'******************************************************************************
'概　要：地域特設ページの検索条件表示関数
'引　数：vCityURL			：地域ページのＵＲＬ　[例]minatoku.asp
'　　　：vCity				：地域名　[例]港区
'　　　：vChikuCode			：地区コード　[例]akasaka→B_Station.Chiku
'　　　：vWorkingTypeCode	：雇用形態コード　[例]002
'　　　：vJobTypeCode		：職種コード　[例]01
'　　　：vChikuSearchFlag	：地区検索表示可否　[0]非表示 [1]表示
'戻り値：
'作成日：2006/12/13
'作成者：Lis Kokubo
'備　考：
'更　新：
'******************************************************************************
Sub DspCitySearchConditionHtml(ByVal vCityURL, ByVal vPrefectureCode, ByVal vCity, ByVal vChikuCode, ByVal vWorkingTypeCode, ByVal vJobTypeCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim idx

	Dim sWTName
	Dim sJTName
	Dim sChikuName
	Dim sCondition1:	sCondition1 = ""
	Dim sCondition2:	sCondition2 = ""

	sWTName = GetDetail("WorkingType", vWorkingTypeCode)
	sJTName = GetJobTypeBig(Left(vJobTypeCode, 2))
	sChikuName = GetChiku(vPrefectureCode, vChikuCode)
	If sWTName <> "" Then
		If sCondition1 <> "" Then sCondition1 = sCondition1 & "、"
		If sCondition2 <> "" Then sCondition2 = sCondition2 & "、"
		sCondition2 = sCondition2 & sWTName
	End If
	If sJTName <> "" Then
		If sCondition1 <> "" Then sCondition1 = sCondition1 & "、"
		If sCondition2 <> "" Then sCondition2 = sCondition2 & "、"
		sCondition1 = sCondition1 & sJTName
	End If
	If sChikuName <> "" Then
		If sCondition1 <> "" Then sCondition1 = sCondition1 & "、"
		If sCondition2 <> "" Then sCondition2 = sCondition2 & "、"
		sCondition1 = sCondition1 & sChikuName
		sCondition2 = sCondition2 & sChikuName
	End If
	If sCondition1 <> "" Then sCondition1 = "&nbsp;&nbsp;<span style=""font-weight:normal;"">()内は、条件「<span style=""font-weight:bold;"">" & sCondition1 & "</span>」での求人票件数。</span>"
	If sCondition2 <> "" Then sCondition2 = "&nbsp;&nbsp;<span style=""font-weight:normal;"">()内は、条件「<span style=""font-weight:bold;"">" & sCondition2 & "</span>」での求人票件数。</span>"

%>
<a name="#search"></a>
<h2 class="ssubtitle">検索条件選択</h2>
<div class="subcontent">
	<table class="citysearch" border="0" cellspacing="0">
		<tbody>
<%
	'*******************************************************************************
	'職種大分類 start
	'*******************************************************************************
%>
		<tr>
			<th class="citysearch" valign="top">
				<p style="float:left; width:60px; margin:0px;">職種</p>
				<p style="float:left; margin:0px;"><%= sCondition2 %></p>
				<p class="citynosearch"><a href="<%= BASEURL %>city/<%= vCityURL %>?wt=<%= vWorkingTypeCode %>&amp;chiku=<%= vChikuCode %>&amp;jt=" title="<%= vCity %>の転職求人情報" style="font-weight:normal;">指定しない</a></p>
				<br clear="all">
			</th>
		</tr>
		<tr>
			<td class="citysearch">
<%
	sSQL = "up_GetListCityJobTypeBig '" & vPrefectureCode & "', '" & vCity & "', '" & vChikuCode & "', '" & vWorkingTypeCode & "', '" & vJobTypeCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	idx = 1
	Do While GetRSState(oRS) = True
		If oRS.Collect("JobTypeCode") = vJobTypeCode Then
			Response.Write "<p class=""citysearch"" style=""color:#ff0000;"">" & oRS.Collect("JobTypeName") & "(" & oRS.Collect("Cnt") & ")</p>"
		ElseIf oRS.Collect("Cnt") > 0 Then
			Response.Write "<p class=""citysearch""><a href=""" & BASEURL & "city/" & vCityURL & "?wt=" & vWorkingTypeCode & "&amp;chiku=" & vChikuCode & "&amp;jt=" & oRS.Collect("JobTypeCode") & """ title=""" & vCity & "の" & oRS.Collect("JobTypeName") & "の転職求人情報"" style=""" & GetConditionStr(oRS.Collect("JobTypeCode"), vJobTypeCode, "color:#ff0000;") & """>" & oRS.Collect("JobTypeName") & "(" & oRS.Collect("Cnt") & ")</a></p>"
		Else
			Response.Write "<p class=""citysearch"">" & oRS.Collect("JobTypeName") & "</p>"
		End If
		If idx Mod 4 = 0 Then Response.Write "<br clear=""all"">"
		oRS.MoveNext
		If GetRSState(oRS) = False And idx Mod 4 <> 0 Then Response.Write "<br clear=""all"">"
		idx = idx + 1
	Loop
	Call RSClose(oRS)
%>
			</td>
		</tr>
<%
	'*******************************************************************************
	'職種大分類 end
	'*******************************************************************************

	'*******************************************************************************
	'雇用形態 start
	'*******************************************************************************
%>
		<tr>
			<th class="citysearch" valign="top">
				<p style="float:left; width:60px; margin:0px;">雇用形態</p>
				<p style="float:left; margin:0px;"><%= sCondition1 %></p>
				<p class="citynosearch"><a href="<%= BASEURL %>city/<%= vCityURL %>?wt=&amp;chiku=<%= vChikuCode %>&amp;jt=<%= vJobTypeCode %>" title="<%= vCity %>の転職求人情報" style="font-weight:normal;">指定しない</a></p>
				<br clear="all">
			</th>
		</tr>
		<tr>
			<td class="citysearch">
<%
	sSQL = "up_GetListCityWorkingType '" & vPrefectureCode & "', '" & vCity & "', '" & vChikuCode & "', '" & vWorkingTypeCode & "', '" & vJobTypeCode & "'"
	Set oRS = QEXE(dbconn, sSQL)
	idx = 1
	Do While GetRSState(oRS) = True
		If oRS.Collect("WorkingTypeCode") = vWorkingTypeCode Then
			Response.Write "<p class=""citysearch"" style=""color:#ff0000;"">" & oRS.Collect("WorkingTypeName") & "(" & oRS.Collect("Cnt") & ")</p>"
		ElseIf oRS.Collect("Cnt") > 0 Then
			Response.Write "<p class=""citysearch""><a href=""" & BASEURL & "city/" & vCityURL & "?wt=" & oRS.Collect("WorkingTypeCode") & "&amp;chiku=" & vChikuCode & "&amp;jt=" & vJobTypeCode & """ title=""" & vCity & "の転職求人情報"" style=""" & GetConditionStr(oRS.Collect("WorkingTypeCode"), vWorkingTypeCode, "color:#ff0000;") & """>" & oRS.Collect("WorkingTypeName") & "(" & oRS.Collect("Cnt") & ")</a></p>"
		Else
			Response.Write "<p class=""citysearch"">" & oRS.Collect("WorkingTypeName") & "</p>"
		End If
		If idx Mod 4 = 0 Then Response.Write "<br clear=""all"">"
		oRS.MoveNext
		If GetRSState(oRS) = False And idx Mod 4 <> 0 Then Response.Write "<br clear=""all"">"
		idx = idx + 1
	Loop
	Call RSClose(oRS)
%>
			</td>
		</tr>
<%
	'*******************************************************************************
	'雇用形態 end
	'*******************************************************************************

	'*******************************************************************************
	'地区 start
	'*******************************************************************************
	If vCity = "港区" Then
%>
		<tr>
			<th class="citysearch" valign="top">
				<p class="citysearch">エリア</p>
				<p class="citynosearch"><a href="<%= BASEURL %>city/<%= vCityURL %>?wt=<%= vWorkingTypeCode %>&amp;chiku=&amp;jt=<%= vJobTypeCode %>" title="<%= vCity %>の転職求人情報" style="font-weight:normal;">指定しない</a></p>
				<br clear="all">
			</th>
		</tr>
		<tr>
			<td class="citysearch">
<%
		sSQL = "up_GetListCityChiku '" & vPrefectureCode & "', '" & vCity & "', '" & vChikuCode & "', '" & vWorkingTypeCode & "', '" & vJobTypeCode & "'"
		Set oRS = QEXE(dbconn, sSQL)
		idx = 1
		Do While GetRSState(oRS) = True
			If oRS.Collect("ChikuCode") = vChikuCode Then
				Response.Write "<p class=""citysearch"" style=""color:#ff0000;"">" & oRS.Collect("ChikuName") & "(" & oRS.Collect("Cnt") & ")</p>"
			ElseIf oRS.Collect("Cnt") > 0 Then
				Response.Write "<p class=""citysearch""><a href=""" & BASEURL & "city/" & vCityURL & "?wt=" & vWorkingTypeCode & "&amp;chiku=" & oRS.Collect("ChikuCode") & "&amp;jt=" & vJobTypeCode & """ title=""" & vCity & "の転職求人情報"" style=""" & GetConditionStr(oRS.Collect("ChikuCode"), vChikuCode, "color:#ff0000;") & """>" & oRS.Collect("ChikuName") & "(" & oRS.Collect("Cnt") & ")</a></p>"
			Else
				Response.Write "<p class=""citysearch"">" & oRS.Collect("ChikuName") & "</p>"
			End If
			If idx Mod 4 = 0 Then Response.Write "<br clear=""all"">"
			oRS.MoveNext
			If GetRSState(oRS) = False And idx Mod 4 <> 0 Then Response.Write "<br clear=""all"">"
			idx = idx + 1
		Loop
		Call RSClose(oRS)
	End If
	'*******************************************************************************
	'地区 end
	'*******************************************************************************
%>
		</tbody>
	</table>
</div>
<%
End Sub

'******************************************************************************
'概　要：vComp1とvComp2が等しい場合は、vStrを取得
'引　数：vComp1	：被比較
'　　　：vComp2	：比較
'　　　：vStr	：vComp1とvComp2が等しい時に返す文字列
'戻り値：
'作成日：2006/12/13
'作成者：Lis Kokubo
'備　考：
'更　新：
'******************************************************************************
Function GetConditionStr(ByVal vComp1, ByVal vComp2, ByVal vStr)
	GetConditionStr = ""
	If vComp1 = vComp2 Then GetConditionStr = vStr
End Function

'******************************************************************************
'概　要：地域特設ページの求人票一覧表示関数
'引　数：vCityURL			：地域ページのＵＲＬ　[例]minatoku.asp
'　　　：vPrefectureCode	：都道府県コード　[例]013
'　　　：vCity				：地域名　[例]港区
'　　　：vChikuCode			：地区コード　[例]akasaka→B_Station.Chiku
'　　　：vWorkingTypeCode	：雇用形態コード　[例]002
'　　　：vJobTypeCode		：職種コード　[例]01
'　　　：vPage				：ページ
'戻り値：
'作成日：2006/12/13
'作成者：Lis Kokubo
'備　考：
'更　新：
'******************************************************************************
Sub DspCityOrderListHtml(ByVal vCityURL, ByVal vPrefectureCode, ByVal vCity, ByVal vChikuCode, ByVal vWorkingTypeCode, ByVal vJobTypeCode, ByVal vPage)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sPageCtrl

	Response.Write "<h2 class=""ssubtitle"">" & vCity& "の転職・求人情報</h2>"
	Response.Write "<div class=""subcontent"">"

	sSQL = "/* 東京都23区人気エリアの求人検索 */"
	sSQL = sSQL & "EXEC up_SearchOrderCity '" & vPrefectureCode & "', '" & vCity & "', '" & vChikuCode & "', '" & vWorkingTypeCode & "', '" & vJobTypeCode & "';"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

	If GetRSState(oRS) = True Then
		sPageCtrl = GetHtmlPageControlParam(dbconn, oRS, 3, vPage, HTTP_CURRENTURL & "city/" & vCityURL & "?wt=" & vWorkingTypeCode & "&chiku=" & vChikuCode & "&jt=" & vJobTypeCode, "")
		Response.Write sPageCtrl
		Response.Write "<div class=""line1"" style=""padding-bottom:5px;""></div>"
		Call DspOrderListDetail3(dbconn, oRS, 3, vPage, "")
		Response.Write "<div class=""line1""></div>"
		Response.Write sPageCtrl
	Else
		Response.Write "<p>お探しの求人情報・お仕事情報はありません。</p>"
	End If

	Response.Write "</div>"
End Sub
%>
