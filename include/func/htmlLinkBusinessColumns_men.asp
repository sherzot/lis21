<%
'*******************************************************************************
'�T�@�v�F�����ƃi�r�r�W�l�X�R���� - �v���[���e�[�V�����J�e�S���̊e�R���������NHTML���擾
'���@���F
'�߂�l�FString
'���@�l�F
'���@���F2010/09/03 LIS K.Kokubo �쐬
'*******************************************************************************
Function htmlLinkBusinessColumns_men(vUserID,vCurrentURL)
	Dim sHTML

	Dim aDate(3)
	Dim idx
	Dim sTitle
	Dim iNo

	iNo = 0

	aDate(0) = CDate("2010/11/22")
	For idx = 1 To UBound(aDate)
		aDate(idx) = aDate(idx-1)+7
	Next


	sTitle = "�t��"
	iNo = iNo + 1
	sHTML = sHTML & "<div style=""text-align:right;""><p class=""m0"">"
	sHTML = sHTML & "&nbsp;��&nbsp;"
	If Date >= aDate(iNo-1) Then
		sHTML = sHTML & "<a href=""/s_contents/businesscolumns/men" & Right("0"&iNo,2) & ".asp"">" & sTitle & "</a>"
		If Date < aDate(iNo-1)+7 Then sHTML = sHTML & "<img src=""/img/new_title/new2.gif"" alt=""NEW!"">"
	Else
		sHTML = sHTML & sTitle
	End If
	sHTML = sHTML & "</p></div>"

	sTitle = "���`�x�[�V����"
	iNo = iNo + 1
	sHTML = sHTML & "<div style=""text-align:right;""><p class=""m0"">"
	sHTML = sHTML & "&nbsp;��&nbsp;"
	If Date >= aDate(iNo-1) Then
		sHTML = sHTML & "<a href=""/s_contents/businesscolumns/men" & Right("0"&iNo,2) & ".asp"">" & sTitle & "</a>"
		If Date < aDate(iNo-1)+7 Then sHTML = sHTML & "<img src=""/img/new_title/new2.gif"" alt=""NEW!"">"
	Else
		sHTML = sHTML & sTitle
	End If
	sHTML = sHTML & "</p></div>"

	sTitle = "�L�����A�A�b�v"
	iNo = iNo + 1
	sHTML = sHTML & "<div style=""text-align:right;""><p class=""m0"">"
	sHTML = sHTML & "&nbsp;��&nbsp;"
	If Date >= aDate(iNo-1) Then
		sHTML = sHTML & "<a href=""/s_contents/businesscolumns/men" & Right("0"&iNo,2) & ".asp"">" & sTitle & "</a>"
		If Date < aDate(iNo-1)+7 Then sHTML = sHTML & "<img src=""/img/new_title/new2.gif"" alt=""NEW!"">"
	Else
		sHTML = sHTML & sTitle
	End If
	sHTML = sHTML & "</p></div>"

	sHTML = sHTML & "<div style=""clear:both;""></div>"


	htmlLinkBusinessColumns_men = sHTML
End Function
%>