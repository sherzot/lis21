<%
'*******************************************************************************
'�T�@�v�F�\�[�V�����u�b�N�}�[�N�ւ̃����N���擾
'���@���F
'�߂�l�FString
'���@�l�F
'���@���F2011/02/24 LIS K.Kokubo �쐬
'*******************************************************************************
Function htmlSocialBookmark(ByVal vURL)
	Dim sHTML

	sHTML = sHTML & "<span class=""link"" style=""cursor:pointer;"" onclick=""sbm_yahoo('" & vURL & "');""><img src=""/img/top/bookmark_yahoo.gif"" width=""16"" height=""16"" alt="""" border=""0""></span>�@"
	sHTML = sHTML & "<span class=""link"" style=""cursor:pointer;"" onclick=""sbm_hatena('" & vURL & "');""><img src=""/img/top/bookmark_b_entry.gif"" border=""0"" alt=""""></span>�@"
	sHTML = sHTML & "<span class=""link"" style=""cursor:pointer;"" onclick=""sbm_buzzurl('" & vURL & "');""><img src=""/img/top/bookmark_b_buzzurl_s.gif"" alt="""" border=""0""></span>�@"
	sHTML = sHTML & "<span class=""link"" style=""cursor:pointer;"" onclick=""sbm_livedoor('" & vURL & "');""><img src=""/img/top/bookmark_livedoor_clip.gif"" width=""16"" height=""16"" alt="""" border=""0""></span>�@"
	sHTML = sHTML & "<span class=""link"" style=""cursor:pointer;"" onclick=""sbm_delicious('" & vURL & "');""><img src=""/img/top/bookmark_delicious_add.gif"" alt=""""></span>"

	htmlSocialBookmark = sHTML
End Function
%>