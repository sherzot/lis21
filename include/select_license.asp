<%
'******************************************************************************
'概　要：【しごとナビ】資格検索のパーツ
'作成者：Lis Kokubo
'作成日：2006/05/11
'備　考：外部JavaScriptファイルにopenerへ値を返す関数を書く。
'　　　：registopener(vidx);
'使用元：company/company_reg3_searchlicense.asp
'******************************************************************************
Dim CONF_idx		: CONF_idx = GetForm("idx", 3)
Dim sGroupCode		: sGroupCode = GetForm("CONF_GroupCode", 1)
Dim sCategoryCode	: sCategoryCode = GetForm("CONF_CategoryCode", 1)

Dim sSelectGroup
Dim sSelectCategoryCode
Dim sSelectCode		: sSelectCode = GetForm("Code", 1)

sSelectGroup = GetLicenseGroupOptionHtml(sGroupCode)
If sGroupCode <> "" Then sSelectCategoryCode = GetLicenseCategoryOptionHtml(sGroupCode, sCategoryCode)
If sCategoryCode <> "" Then sSelectCode = GetLicenseOptionHtml(sGroupCode, sCategoryCode, "")
%>
<input type="hidden" name="idx" value="<%= CONF_idx %>">
<table width="600px">
	<tr>
		<td bgcolor="#339933"><font color="#ffffff">資格選択</font></td>
	</tr>
	<tr>
		<td style="padding-top:5px;padding-left:1em;border-left:solid 1px #339933;border-right:solid 1px #339933;">
			<select onchange="document.forms.idfrmPostBack.submit();" name="CONF_GroupCode" style="width:150px;">
				<option value="">------大分類------</option>
				<%= sSelectGroup %>
			</select>
		</td>
	</tr>
	<tr>
		<td style="padding-left:1.5em;border-left:solid 1px #339933;border-right:solid 1px #339933;">
			<span>┗</span>
			<select onchange="document.forms.idfrmPostBack.submit();" name="CONF_CategoryCode" style="width:200px;">
				<option value="">---------中分類----------</option>
				<%= sSelectCategoryCode %>
			</select>
		</td>
	</tr>
	<tr>
		<td style="padding-left:3em;padding-bottom:5px;border-left:solid 1px #339933;border-right:solid 1px #339933;border-bottom:solid 1px #339933;">
			<span>┗</span>
			<select name="CONF_Code" style="width:410px;">
				<option value="">-----------------------小分類-------------------------</option>
				<%= sSelectCode %>
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
