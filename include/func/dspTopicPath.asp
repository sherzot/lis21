<%
'*******************************************************************************
'�T�@�v�F�p���������X�g�\��
'���@���F
'�@�@�@�F
'�߂�l�F
'�쐬���F2007/03/08 Lis Kokubo
'���@���F
'*******************************************************************************
Function DspTopicPath(ByVal vUserType, ByVal vName1, ByVal vURL1, ByVal vName2, ByVal vURL2, ByVal vName3, ByVal vURL3, ByVal vName4, ByVal vURL4, ByVal vName5, ByVal vURL5, ByVal vName6, ByVal vURL6, ByVal vName7, ByVal vURL7, ByVal vName8, ByVal vURL8, ByVal vName9, ByVal vURL9, ByVal vName10, ByVal vURL10)
	DspTopicPath = ""
	Call GetPartTopicPath(DspTopicPath, "�����ƃi�r", BASEURL)
	If vUserType = "staff" Then Call GetPartTopicPath(DspTopicPath, "My&nbsp;�y�[�W", BASEURL & "staff/s_login.asp")
	If vUserType = "company" Then Call GetPartTopicPath(DspTopicPath, "My&nbsp;�y�[�W", BASEURL & "company/c_login.asp")
	'If vUserType = "dispatch" Then Call GetPartTopicPath(DspTopicPath, "My&nbsp;Page", BASEURL & "dispatch/d_login.asp")

	If G_FLGRESUME = False Then
		Call GetPartTopicPath(DspTopicPath, vName1, vURL1)
		Call GetPartTopicPath(DspTopicPath, vName2, vURL2)
		Call GetPartTopicPath(DspTopicPath, vName3, vURL3)
		Call GetPartTopicPath(DspTopicPath, vName4, vURL4)
		Call GetPartTopicPath(DspTopicPath, vName5, vURL5)
		Call GetPartTopicPath(DspTopicPath, vName6, vURL6)
		Call GetPartTopicPath(DspTopicPath, vName7, vURL7)
		Call GetPartTopicPath(DspTopicPath, vName8, vURL8)
		Call GetPartTopicPath(DspTopicPath, vName9, vURL9)
		Call GetPartTopicPath(DspTopicPath, vName10, vURL10)
	End If
End Function

'*******************************************************************************
'�T�@�v�F�p���������X�g�̌X�̃A���J�[�^�O���擾
'���@���FrPath
'�@�@�@�FvName	�F�����N��
'�@�@�@�FvURL	�F
'�߂�l�F
'�쐬���F2007/03/08 Lis Kokubo
'���@���F
'*******************************************************************************
Function GetPartTopicPath(ByRef rPath, ByVal vName, ByVal vURL)
	If vName & vURL <> "" Then
		If rPath <> "" Then rPath = rPath & "<span style=""padding-left:5px;padding-right:5px;"">&gt;</span>"
		If vURL <> "" Then rPath = rPath & "<a href=""" & vURL & """ title=""" & vName & """>"
		If vName <> "" Then rPath = rPath & vName
		If vURL <> "" Then rPath = rPath & "</a>"
	End If
End Function
%>
