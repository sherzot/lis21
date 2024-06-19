<%
'******************************************************************************
'概　要：【しごとナビ】職種検索のパーツ
'作成者：Lis Kokubo
'作成日：2006/05/02
'備　考：外部JavaScriptファイルにopenerへ値を返す関数を書く。
'　　　：registopener(vidx);
'使用元：company/mailhistory_search_jobtype_company.asp
'　　　：company/company_reg2_searchjobtype.asp
'******************************************************************************
Dim CONF_idx				: CONF_idx = GetForm("idx", 3)
Dim CONF_BigClassCode		: CONF_BigClassCode = GetForm("CONF_BigClassCode", 1)
Dim CONF_BigClassName		: CONF_BigClassName = GetForm("CONF_BigClassName", 1)
Dim CONF_MiddleClassCode	: CONF_MiddleClassCode = GetForm("CONF_MiddleClassCode", 1)
Dim CONF_MiddleClassName	: CONF_MiddleClassName = GetForm("CONF_MiddleClassName", 1)
Dim CONF_JobTypeCode		: CONF_JobTypeCode = GetForm("CONF_JobTypeCode", 1)

Dim sBigClassOption
Dim sMiddleClassOption
Dim sJobTypeName

sBigClassOption = GetJobTypeBigClassOptionHtml(CONF_BigClassCode)
If CONF_BigClassCode <> "" Then sMiddleClassOption = GetJobTypeOptionHtml(CONF_BigClassCode, CONF_MiddleClassCode)
%>
<input type="hidden" name="idx" value="<%= CONF_idx %>">
<div align="center">
<table width="500">
	<tr>
		<td colspan="2" bgcolor="#339933"><font color="#ffffff">職種選択</font></td>
	</tr>
	<tr bgcolor="#ccff99" class="moji912">
		<td align="left">(1)大分類を選択して下さい</td>
		<td align="left">(2)中分類を選択して下さい</td>
	</tr>
	<tr>
		<td>
			<select id="idsltBigClassCode" name="CONF_BigClassCode" style="width:150px;" onchange="document.getElementById('idfrmPostBack').submit();">
				<option value=""></option>
				<%= sBigClassOption %>
			</select>
		</td>
		<td>
			<select id="idsltJobTypeCode" style="width:400px;" name="CONF_JobTypeCode">
				<option value=""></option>
				<%= sMiddleClassOption %>
			</select>
			<input id="idJobTypeName" name="CONF_JobTypeName" type="hidden" value="">
		</td>
	</tr>
	<tr>
		<td colspan="2" align="right">
			<input type="button" name="ok" value="(3)決定" onclick="registopener('<%= CONF_idx %>');">
			<input type="button" name="cancel" value="キャンセル" onclick="window.close();">
		</td>
	</tr>
</table>
</div>
