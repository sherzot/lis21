<%
'*******************************************************************************
'�T�@�v�FHTML������DOCTYPE�`body�^�O�܂ł��擾
'���@���FvSite			�F�T�C�g�̃��[�g�t�q�k �i��Fhttp://www.shigotonavi.co.jp/)
'�@�@�@�FvTitle			�F�y�[�W�^�C�g��
'�@�@�@�FvKeywords		�F�y�[�W�L�[���[�h
'�@�@�@�FvDescription	�F�y�[�W������
'�@�@�@�FvAddHead		�F<head></head>�̒��Ɋ܂߂郁�^ (��F<link>�^�O<script>�^�O�Ȃǂ̊O���t�@�C����`�Ȃ�)
'�@�@�@�FvIndexFlag		�F�N���[���[���y�[�W��o�^���邱�Ƃ̉ۃt���O [True]���� [<>True]�s��
'�@�@�@�FvFollowFlag	�F�N���[���[���y�[�W�̃����N�����ǂ邱�Ƃ̉ۃt���O [True]���� [<>True]�s��
'�@�@�@�FvArchiveFlag	�F�N���[���[���y�[�W�L���b�V�����邱�Ƃ̉ۃt���O [True]���� [<>True]�s��
'�@�@�@�FvCacheFlag		�F���[�U�̂o�b�Ƀy�[�W���L���b�V�����邱�Ƃ̉ۃt���O [True]���� [<>True]�s��
'�@�@�@�FvBodyAttribute	�F<body>�̑���
'�o�@�́F
'�߂�l�FString
'���@�l�F
'���@���F2010/05/11 LIS K.Kokubo �쐬
'	   �F2012/06/18 LIS.T.Seki �ҏW�@</head>�y��<body>�̔r��
'�@�@�@�F2016/10/24 LIS Y.yamasaki SEO�w�����̓K�p
'�@�@�@�F2017/07/20 LIS SEO�w�����̓K�p
'*******************************************************************************
Function htmlHeader(ByVal vSite, ByVal vTitle, ByVal vKeywords, ByVal vDescription, ByVal vAddHead, ByVal vIndexFlag, ByVal vFollowFlag, ByVal vArchiveFlag, ByVal vCacheFlag, ByVal vBodyAttribute)
	Dim sHTML
	Dim sRobots
	Dim sCache
	Dim sContentType
        Dim FULLURL
	sHTML = ""
	sRobots = ""
	sCache = ""

	'20170720 �A�N�Z�X�����y�[�W��HTTPS��HTTP���ɂ���āAcanonical�^�O����URL��ύX����
	if Request.ServerVariables("HTTPS") = "off" then
        	FULLURL = "http://www.shigotonavi.co.jp" & Request.ServerVariables("URL")
	else
        	FULLURL = "https://www.shigotonavi.co.jp" & Request.ServerVariables("URL")
	end if
		
	If vIndexFlag = True Then: sRobots = sRobots & "index": Else: sRobots = sRobots & "noindex": End If
	If vFollowFlag = True Then: sRobots = sRobots & ",follow": Else: sRobots = sRobots & ",nofollow": End If
	If vArchiveFlag = True Then: sRobots = sRobots & ",archive": Else: sRobots = sRobots & ",noarchive": End If
	If vCacheFlag = True Then: sCache = sCache & "cache": Else: sCache = sCache & "nocache": End If

	sHTML = sHTML & "<!DOCTYPE HTML>"
	sHTML = sHTML & "<html>"
	
	sHTML = sHTML & "<head>"
	sHTML = sHTML & "<meta http-equiv=""X-UA-Compatible"" content=""IE=8 ; IE=9"" />"
	
	sHTML = sHTML & "<meta name=""viewport"" content=""width=device-width,user-scalable=no,maximum-scale=1"">"

        sHTML = sHTML & "<link rel=""canonical"" href=""" & FULLURL & """ />"

	sHTML = sHTML & "<meta http-equiv=""content-type"" content=""text/html; charset=shift_jis"">"
	sHTML = sHTML & "<meta http-equiv=""content-script-type"" content=""text/javascript"">"
	sHTML = sHTML & "<meta http-equiv=""content-style-type"" content=""text/css"">"
	sHTML = sHTML & "<meta name=""robots"" content=""" & sRobots & """>"
	sHTML = sHTML & "<meta name=""googlebot"" content=""" & sRobots & """>"
	sHTML = sHTML & "<meta name=""keywords"" content=""" & vKeywords & """>"
	sHTML = sHTML & "<meta name=""description"" content=""" & vDescription & """>"
	sHTML = sHTML & "<![if !IE]><script src=""/script/jquery.js"" type=""text/javascript""></script><![endif]>"
	sHTML = sHTML & "<!--[if gte IE 9]><script src=""/script/jquery.js"" type=""text/javascript""></script><![endif]-->"
	sHTML = sHTML & "<!--[if lt IE 9]><script src=""/script/jquery_for_ie.js""></script><![endif]-->"
	
	sHTML = sHTML & "<script src=""/script/base.js"" type=""text/javascript""></script>"
	sHTML = sHTML & "<![if !IE]><script src=""/script/top10.js"" type=""text/javascript""></script><![endif]>"
	sHTML = sHTML & "<!--[if gte IE 9]><script src=""/script/top10.js"" type=""text/javascript""></script><![endif]-->"
	sHTML = sHTML & "<!--[if lt IE 9]><script src=""/script/top10_for_ie.js""></script><![endif]-->"
	sHTML = sHTML & "<!--[if lt IE 9]><script src=""/script/html5.js""></script><![endif]-->"
	sHTML = sHTML & "<link rel=""stylesheet"" type=""text/css"" href=""/css/c_company_main.css"">"
	
	sHTML = sHTML & "<link rel=""stylesheet"" type=""text/css"" href=""/css/smartphone/base.css"">"
	sHTML = sHTML & "<link rel=""stylesheet"" type=""text/css"" href=""/css/smartphone/header_smart.css"">"
	sHTML = sHTML & "<link rel=""stylesheet"" type=""text/css"" href=""/css/smartphone/footer_smart.css"">"
	sHTML = sHTML & "<script src=""/script/smart/smapho.js"" type=""text/javascript""></script>"
	
	sHTML = sHTML & "<link rel=""apple-touch-icon-precomposed"" href=""logoIcon.png"" />"
	
	sHTML = sHTML & vAddHead
	sHTML = sHTML & "<title>" & vTitle & "</title>"
	
	htmlHeader = sHTML
End Function
%>

<%
'*******************************************************************************
'�T�@�v�FHTML������DOCTYPE�`body�^�O�܂ł��擾
'���@���FvSite			�F�T�C�g�̃��[�g�t�q�k �i��Fhttp://www.shigotonavi.co.jp/)
'�@�@�@�FvTitle			�F�y�[�W�^�C�g��
'�@�@�@�FvKeywords		�F�y�[�W�L�[���[�h
'�@�@�@�FvDescription	�F�y�[�W������
'�@�@�@�FvAddHead		�F<head></head>�̒��Ɋ܂߂郁�^ (��F<link>�^�O<script>�^�O�Ȃǂ̊O���t�@�C����`�Ȃ�)
'�@�@�@�FvIndexFlag		�F�N���[���[���y�[�W��o�^���邱�Ƃ̉ۃt���O [True]���� [<>True]�s��
'�@�@�@�FvFollowFlag	�F�N���[���[���y�[�W�̃����N�����ǂ邱�Ƃ̉ۃt���O [True]���� [<>True]�s��
'�@�@�@�FvArchiveFlag	�F�N���[���[���y�[�W�L���b�V�����邱�Ƃ̉ۃt���O [True]���� [<>True]�s��
'�@�@�@�FvCacheFlag		�F���[�U�̂o�b�Ƀy�[�W���L���b�V�����邱�Ƃ̉ۃt���O [True]���� [<>True]�s��
'�@�@�@�FvBodyAttribute	�F<body>�̑���
'�@�@�@�FvCanonical	�F�J�m�j�J���^�O
'�o�@�́F
'�߂�l�FString
'���@�l�F
'���@���F2010/05/11 LIS K.Kokubo �쐬
'	   �F2012/06/18 LIS.T.Seki �ҏW�@</head>�y��<body>�̔r��
'�@�@�@�F2016/10/24 LIS Y.yamasaki �r�d�n�w�����̓K�p
'�@�@�@�F2017/04/26 K.K �J�m�j�J���^�O�ʐݒ�Ή�
'*******************************************************************************
Function htmlHeader_Canonical(ByVal vSite, ByVal vTitle, ByVal vKeywords, ByVal vDescription, ByVal vAddHead, ByVal vIndexFlag, ByVal vFollowFlag, ByVal vArchiveFlag, ByVal vCacheFlag, ByVal vBodyAttribute, ByVal vCanonical)
	Dim sHTML
	Dim sRobots
	Dim sCache
	Dim sContentType
        Dim FULLURL
	sHTML = ""
	sRobots = ""
	sCache = ""
        FULLURL = "http://www.shigotonavi.co.jp" & Request.ServerVariables("URL")
	If vIndexFlag = True Then: sRobots = sRobots & "index": Else: sRobots = sRobots & "noindex": End If
	If vFollowFlag = True Then: sRobots = sRobots & ",follow": Else: sRobots = sRobots & ",nofollow": End If
	If vArchiveFlag = True Then: sRobots = sRobots & ",archive": Else: sRobots = sRobots & ",noarchive": End If
	If vCacheFlag = True Then: sCache = sCache & "cache": Else: sCache = sCache & "nocache": End If

	sHTML = sHTML & "<!DOCTYPE HTML>"
	sHTML = sHTML & "<html>"
	
	sHTML = sHTML & "<head>"
	sHTML = sHTML & "<meta http-equiv=""X-UA-Compatible"" content=""IE=8 ; IE=9"" />"
	
	sHTML = sHTML & "<meta name=""viewport"" content=""width=device-width,user-scalable=no,maximum-scale=1"">"

        sHTML = sHTML & "<link rel=""canonical"" href=""" & vCanonical & """ />"

	sHTML = sHTML & "<meta http-equiv=""content-type"" content=""text/html; charset=shift_jis"">"
	sHTML = sHTML & "<meta http-equiv=""content-script-type"" content=""text/javascript"">"
	sHTML = sHTML & "<meta http-equiv=""content-style-type"" content=""text/css"">"
	sHTML = sHTML & "<meta name=""robots"" content=""" & sRobots & """>"
	sHTML = sHTML & "<meta name=""googlebot"" content=""" & sRobots & """>"
	sHTML = sHTML & "<meta name=""keywords"" content=""" & vKeywords & """>"
	sHTML = sHTML & "<meta name=""description"" content=""" & vDescription & """>"
	sHTML = sHTML & "<![if !IE]><script src=""/script/jquery.js"" type=""text/javascript""></script><![endif]>"
	sHTML = sHTML & "<!--[if gte IE 9]><script src=""/script/jquery.js"" type=""text/javascript""></script><![endif]-->"
	sHTML = sHTML & "<!--[if lt IE 9]><script src=""/script/jquery_for_ie.js""></script><![endif]-->"
	
	sHTML = sHTML & "<script src=""/script/base.js"" type=""text/javascript""></script>"
	sHTML = sHTML & "<![if !IE]><script src=""/script/top10.js"" type=""text/javascript""></script><![endif]>"
	sHTML = sHTML & "<!--[if gte IE 9]><script src=""/script/top10.js"" type=""text/javascript""></script><![endif]-->"
	sHTML = sHTML & "<!--[if lt IE 9]><script src=""/script/top10_for_ie.js""></script><![endif]-->"
	sHTML = sHTML & "<!--[if lt IE 9]><script src=""/script/html5.js""></script><![endif]-->"
	sHTML = sHTML & "<link rel=""stylesheet"" type=""text/css"" href=""/css/c_company_main.css"">"
	
	sHTML = sHTML & "<link rel=""stylesheet"" type=""text/css"" href=""/css/smartphone/base.css"">"
	sHTML = sHTML & "<link rel=""stylesheet"" type=""text/css"" href=""/css/smartphone/header_smart.css"">"
	sHTML = sHTML & "<link rel=""stylesheet"" type=""text/css"" href=""/css/smartphone/footer_smart.css"">"
	sHTML = sHTML & "<script src=""/script/smart/smapho.js"" type=""text/javascript""></script>"
	
	sHTML = sHTML & "<link rel=""apple-touch-icon-precomposed"" href=""logoIcon.png"" />"
	
	sHTML = sHTML & vAddHead
	sHTML = sHTML & "<title>" & vTitle & "</title>"
	
	htmlHeader_Canonical = sHTML
End Function
%>