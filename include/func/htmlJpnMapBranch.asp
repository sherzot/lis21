<%
'*******************************************************************************
'概　要：日本地図上の拠点をプロットしたイメージのHTMLを取得
'引　数：
'出　力：
'戻り値：String
'備　考：
'履　歴：2011/04/21 LIS K.Kokubo 作成
'*******************************************************************************
Function htmlJpnMapBranch()
	Dim sHTML

	sHTML = "<div style=""text-align:center;"">" & _
		"<img src=""/img/common/jpnmapbranch.gif"" alt=""リスのオフィス所在地"" border=""0"" usemap=""#jpnmapbranch"">" & _
		"<map name=""jpnmapbranch"">" & _
		"<area shape=""rect"" coords=""377,119,463,133"" alt=""群馬オフィス"" href=""" & HTTP_LIS_CURRENTURL & "lis/office_map.asp?id=2"" target=""_blank"">" & _
		"<area shape=""rect"" coords=""404,147,462,160"" alt=""東京本社"" href=""" & HTTP_LIS_CURRENTURL & "lis/office_map.asp?id=1"" target=""_blank"">" & _
		"<area shape=""rect"" coords=""404,175,462,189"" alt=""静岡支社"" href=""" & HTTP_LIS_CURRENTURL & "lis/office_map.asp?id=3"" target=""_blank"">" & _
		"<area shape=""rect"" coords=""390,202,462,216"" alt=""名古屋支社"" href=""" & HTTP_LIS_CURRENTURL & "lis/office_map.asp?id=4"" target=""_blank"">" & _
		"<area shape=""rect"" coords=""404,226,462,240"" alt=""大阪支社"" href=""" & HTTP_LIS_CURRENTURL & "lis/office_map.asp?id=5"" target=""_blank"">" & _
		"<area shape=""rect"" coords=""29,110,115,124"" alt=""岡山オフィス"" href=""" & HTTP_LIS_CURRENTURL & "lis/office_map.asp?id=18"" target=""_blank"">" & _
		"<area shape=""rect"" coords=""30,145,88,159"" alt=""広島支社"" href=""" & HTTP_LIS_CURRENTURL & "lis/office_map.asp?id=6"" target=""_blank"">" & _
		"</map>" & _
		"</div>"

	htmlJpnMapBranch = sHTML
End Function
%>
