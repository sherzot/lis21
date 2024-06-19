<%
'******************************************************************************
'概　要：【しごとナビ】駅検索のパーツ
'作成者：Lis Kokubo
'作成日：2006/05/02
'備　考：
'使用元：company/mailhistory_search_jobtype_company.asp
'　　　：company/company_reg2_searchjobtype.asp
'******************************************************************************
Dim CONF_idx			: CONF_idx = GetForm("idx", 3)
Dim sPrefectureCode		: sPrefectureCode = GetForm("PrefectureCode", 1)

Dim sSelectPrefecture
Dim sSelectRailwayLine

sSelectPrefecture = GetPrefectureOptionHtml(sPrefectureCode)
If sPrefectureCode <> "" Then sSelectRailwayLine = GetRailwayLineOptionHtml(sPrefectureCode, "")
%>
<input type="hidden" name="idx" value="<%= CONF_idx %>">
<table width="500px">
	<tr>
		<td colspan="2" bgcolor="#339933"><font color="#ffffff">駅選択</font></td>
	</tr>
	<tr bgcolor="#ccff99" class="moji912">
		<td align="left">(1)都道府県</td>
		<td align="left">(2)路線</td>
	</tr>
	<tr>
		<td width="100px">
			<select onchange="document.forms.idfrmPostBack.submit();" name="PrefectureCode" style="width: 100px;">
				<option value=""></option>
				<%= sSelectPrefecture %>
			</select>
		</td>
		<td width="400px">
			<select name="RailwayLineCode" style="width: 400px;">
				<option value=""></option>
				<%= sSelectRailwayLine %>
			</select>
		</td>
	</tr>
	<tr>
		<td colspan="3" align="right">
			<input type="button" name="ok" value="決定" onclick="registopener('<%= CONF_idx %>');">
			<input type="button" name="cancel" value="キャンセル" onclick="window.close();">
		</td>
	</tr>
</table>
