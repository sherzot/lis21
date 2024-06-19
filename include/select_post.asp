<%
'******************************************************************************
'概　要：【しごとナビ】郵便番号検索のパーツ
'作成者：Lis Kokubo
'作成日：2006/05/02
'備　考：外部JavaScriptファイルにopenerへ値を返す関数を書く。
'　　　：registopener();
'使用元：staff/peron_reg1.asp
'******************************************************************************
Dim sSQL_select_post
Dim oRS_select_post
Dim CONF_Post_U	: CONF_Post_U = GetForm("u", 2)
Dim CONF_Post_L	: CONF_Post_L = GetForm("l", 2)
Dim sPost_U
Dim sPost_L
Dim sPrefectureCode
Dim sPrefectureName
Dim sPrefectureName_F
Dim sCity
Dim sCity_F
Dim sTown
Dim sTown_F

sSQL_select_post = "up_DtlZip '" & CONF_Post_U & "', '" & CONF_Post_L & "'"
Set oRS_select_post = dbconn.Execute(sSQL_select_post)
If GetRSState(oRS_select_post) = True Then
	sPost_U = oRS_select_post.Fields("Post_U").Value
	sPost_L = oRS_select_post.Fields("Post_L").Value
	sPrefectureCode = oRS_select_post.Fields("PrefectureCode").Value
	sPrefectureName = oRS_select_post.Fields("PrefectureName").Value
	sPrefectureName_F = oRS_select_post.Fields("PrefectureName_F").Value
	sCity = oRS_select_post.Fields("City").Value
	sCity_F = oRS_select_post.Fields("City_F").Value
	sTown = oRS_select_post.Fields("Town").Value
	sTown_F = oRS_select_post.Fields("Town_F").Value
	oRS_select_post.Close
End If
%>
<input type="hidden" name="CONF_Post_U" value="<%= sPost_U %>">
<input type="hidden" name="CONF_Post_L" value="<%= sPost_L %>">
<input type="hidden" name="CONF_PrefectureCode" value="<%= sPrefectureCode %>">
<input type="hidden" name="CONF_PrefectureName" value="<%= sPrefectureName %>">
<input type="hidden" name="CONF_PrefectureName_F" value="<%= sPrefectureName_F %>">
<input type="hidden" name="CONF_City" value="<%= sCity %>">
<input type="hidden" name="CONF_City_F" value="<%= sCity_F %>">
<input type="hidden" name="CONF_Town" value="<%= sTown %>">
<input type="hidden" name="CONF_Town_F" value="<%= sTown_F %>">
