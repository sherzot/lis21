<%
'*******************************************************************************
'�T�@�v�F�^�uIndex���擾
'���@���F
'�߂�l�FString
'���@�l�F
'���@���F2010/05/13 LIS K.Kokubo �쐬
'*******************************************************************************
Function htmlTabIndex(ByVal vURL,ByVal vUserType,ByVal vHeadComment)
	Dim sHTML
	Dim aPattern
	Dim aLink(5)
	Dim iTabIndexType
	Dim sBorderColor

	aPattern = Array(1,1,1,1,1,1,1,1)
	vURL = LCase(vURL)

	iTabIndexType = getTabIndexType(vURL)
	sBorderColor = "#66cc33"
	If iTabIndexType = 5 Or iTabIndexType = 7 Then sBorderColor = "#0033cc"

	Select Case iTabIndexType
	'�͂��߂Ă̕��^�u
	Case 0: aPattern(0) = 2
	'���l��T���^�u
	Case 1: aPattern(1) = 2
	'�֗��c�[���^�u
	Case 2: aPattern(2) = 2
	'�]�E�T�|�[�g�^�u
	Case 3: aPattern(3) = 2
	'�R�~���j�e�B�^�u
	Case 4: aPattern(4) = 2
	'�̗p���S���҃^�u
	Case 5: aPattern(5) = 2
	'�o�^����^�u
	Case 6: aPattern(6) = 2
	'���O�C����ƃ^�u
	Case 7: aPattern(7) = 2
	'Case Else: aPattern(1) = 2
	End Select

'	If vUserType = "staff" Then
'		aLink(0) = "<a href=""" & HTTPS_CURRENTURL & "tab/index6.asp"" title=""�悤�����I�@�����ƃi�r�ցI�@���Ȃ��̂��߂̃��j���[�ꗗ�ł��B""><img src=""/img/common/tab_index7_" & aPattern(6) & ".png"" alt=""�悤�����I�@�����ƃi�r�ցI�@���Ȃ��̂��߂̃��j���[�ꗗ�ł��B"" border=""0""></a>"
'	ElseIf vUserType = "company" Then
'		aLink(0) = "<a href=""" & HTTPS_CURRENTURL & "tab/index7.asp"" title=""�悤�����I�@�����ƃi�r�ցI�@���Ȃ��̂��߂̃��j���[�ꗗ�ł��B""><img src=""/img/common/tab_index8_" & aPattern(7) & ".png"" alt=""�悤�����I�@�����ƃi�r�ցI�@���Ȃ��̂��߂̃��j���[�ꗗ�ł��B"" border=""0""></a>"
'	Else
'		aLink(0) = "<a href=""" & HTTPS_CURRENTURL & "tab/index1.asp"" title=""����o�^�i�������o�^�j���郁���b�g�Ƃ́H""><img src=""/img/common/tab_index1_" & aPattern(0) & ".png"" alt=""����o�^�i�������o�^�j���郁���b�g�Ƃ́H"" border=""0""></a>"
'	End If

	aLink(0) = "<a href=""" & HTTPS_CURRENTURL & """>HOME</a>"
	aLink(1) = "<a href=""" & HTTPS_CURRENTURL & "search/"">�����Ƃ�T��</a>"	
	'aLink(2) = "<a href=""" & HTTPS_CURRENTURL & "koryu/"">��</a>"
	aLink(3) = "<a href=""" & HTTPS_CURRENTURL & "manabu/"">�w��</a>"
	aLink(4) = "<a href=""" & HTTPS_CURRENTURL & "link/"">�����N</a>"
	aLink(5) = "<a href=""" & HTTPS_CURRENTURL & "company_AB/"">�̗p���</a>"
	


	sHTML = ""
	sHTML = sHTML & "<nav class=""headtab"">"
	sHTML = sHTML & "<div id=""line_on"">"	
	sHTML = sHTML & "<ul>"
	sHTML = sHTML & "<li id=""gmenu01"">" & aLink(1) & "</li>"
	'sHTML = sHTML & "<li id=""gmenu02"">" & aLink(2) & "</li>"
	sHTML = sHTML & "<li id=""gmenu03"">" & aLink(3) & "</li>"
	sHTML = sHTML & "<li id=""gmenu04"">" & aLink(4) & "</li>"
	sHTML = sHTML & "<li id=""gmenu05"">" & aLink(5) & "</li>"
	sHTML = sHTML & "<li id=""gmenu06"">" & aLink(0) & "</li>"
	sHTML = sHTML & "</ul>"
	sHTML = sHTML & "<img src=""/img/border.gif"" id=""slide_border""></div>"
	
	sHTML = sHTML & "</nav>"

	htmlTabIndex = sHTML
End Function
%>
