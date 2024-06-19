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
Dim sRailwayLineCode	: sRailwayLineCode = GetForm("RailwayLineCode", 1)

Dim sSelectPrefecture
Dim sSelectRailwayLine
Dim sSelectStation		: sSelectStation = GetForm("StationCode", 1)

sSelectPrefecture = GetPrefectureOptionHtml(sPrefectureCode)
If sPrefectureCode <> "" Then sSelectRailwayLine = GetRailwayLineOptionHtml(sPrefectureCode, sRailwayLineCode)
If sRailwayLineCode <> "" Then sSelectStation = GetStationOptionHtml(sPrefectureCode, sRailwayLineCode, "")
%>
<input type="hidden" name="idx" value="<%= CONF_idx %>">
<table width="600px">
	<tr>
		<td colspan="3" bgcolor="#339933"><font color="#ffffff">駅選択</font></td>
	</tr>
	<tr bgcolor="#ccff99" class="moji912">
		<td align="left">(1)都道府県</td>
		<td align="left">(2)路線</td>
		<td align="left">(3)駅</td>
	</tr>
	<tr>
		<td width="100px">
			<select onchange="document.forms.idfrmPostBack.submit();" name="PrefectureCode" style="width: 100px;">
				<option value=""></option>
				<%= sSelectPrefecture %>
			</select>
		</td>
		<td width="300px">
			<select onchange="document.forms.idfrmPostBack.submit();" name="RailwayLineCode" style="width: 300px;">
				<option value=""></option>
				<%= sSelectRailwayLine %>
			</select>
		</td>
		<td widt="200px">
			<select name="StationCode" style="width: 200px;">
				<option value=""></option>
				<%= sSelectStation %>
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
