<%
'*******************************************************************************
'�T�@�v�F���{�n�}��̋��_���v���b�g�����C���[�W��HTML���擾
'���@���F
'�o�@�́F
'�߂�l�FString
'���@�l�F
'���@���F2011/04/21 LIS K.Kokubo �쐬
'*******************************************************************************
Function htmlJpnMapBranch()
	Dim sHTML

	sHTML = "<div style=""text-align:center;"">" & _
		"<img src=""/img/common/jpnmapbranch.gif"" alt=""���X�̃I�t�B�X���ݒn"" border=""0"" usemap=""#jpnmapbranch"">" & _
		"<map name=""jpnmapbranch"">" & _
		"<area shape=""rect"" coords=""377,119,463,133"" alt=""�Q�n�I�t�B�X"" href=""" & HTTP_LIS_CURRENTURL & "lis/office_map.asp?id=2"" target=""_blank"">" & _
		"<area shape=""rect"" coords=""404,147,462,160"" alt=""�����{��"" href=""" & HTTP_LIS_CURRENTURL & "lis/office_map.asp?id=1"" target=""_blank"">" & _
		"<area shape=""rect"" coords=""404,175,462,189"" alt=""�É��x��"" href=""" & HTTP_LIS_CURRENTURL & "lis/office_map.asp?id=3"" target=""_blank"">" & _
		"<area shape=""rect"" coords=""390,202,462,216"" alt=""���É��x��"" href=""" & HTTP_LIS_CURRENTURL & "lis/office_map.asp?id=4"" target=""_blank"">" & _
		"<area shape=""rect"" coords=""404,226,462,240"" alt=""���x��"" href=""" & HTTP_LIS_CURRENTURL & "lis/office_map.asp?id=5"" target=""_blank"">" & _
		"<area shape=""rect"" coords=""29,110,115,124"" alt=""���R�I�t�B�X"" href=""" & HTTP_LIS_CURRENTURL & "lis/office_map.asp?id=18"" target=""_blank"">" & _
		"<area shape=""rect"" coords=""30,145,88,159"" alt=""�L���x��"" href=""" & HTTP_LIS_CURRENTURL & "lis/office_map.asp?id=6"" target=""_blank"">" & _
		"</map>" & _
		"</div>"

	htmlJpnMapBranch = sHTML
End Function
%>
