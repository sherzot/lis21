<%
'*******************************************************************************
'概　要：HTML文書のDOCTYPE～bodyタグまでを取得
'引　数：vSite			：サイトのルートＵＲＬ （例：http://www.shigotonavi.co.jp/)
'　　　：vTitle			：ページタイトル
'　　　：vKeywords		：ページキーワード
'　　　：vDescription	：ページ説明文
'　　　：vAddHead		：<head></head>の中に含めるメタ (例：<link>タグ<script>タグなどの外部ファイル定義など)
'　　　：vIndexFlag		：クローラーがページを登録することの可否フラグ [True]許可 [<>True]不可
'　　　：vFollowFlag	：クローラーがページのリンクをたどることの可否フラグ [True]許可 [<>True]不可
'　　　：vArchiveFlag	：クローラーがページキャッシュすることの可否フラグ [True]許可 [<>True]不可
'　　　：vCacheFlag		：ユーザのＰＣにページをキャッシュすることの可否フラグ [True]許可 [<>True]不可
'　　　：vBodyAttribute	：<body>の属性
'出　力：
'戻り値：String
'備　考：
'履　歴：2010/05/11 LIS K.Kokubo 作成
'	   ：2012/06/18 LIS.T.Seki 編集　</head>及び<body>の排除
'      ：2015/09/01 Lis K.Kimura ピンチできるようにViewportを書き換え
'*******************************************************************************
Function htmlHeader2Pinch(ByVal vSite, ByVal vTitle, ByVal vKeywords, ByVal vDescription, ByVal vAddHead, ByVal vIndexFlag, ByVal vFollowFlag, ByVal vArchiveFlag, ByVal vCacheFlag, ByVal vBodyAttribute)
	Dim sHTML
	Dim sRobots
	Dim sCache
	Dim sContentType

	sHTML = ""
	sRobots = ""
	sCache = ""

	If vIndexFlag = True Then: sRobots = sRobots & "index": Else: sRobots = sRobots & "noindex": End If
	If vFollowFlag = True Then: sRobots = sRobots & ",follow": Else: sRobots = sRobots & ",nofollow": End If
	If vArchiveFlag = True Then: sRobots = sRobots & ",archive": Else: sRobots = sRobots & ",noarchive": End If
	If vCacheFlag = True Then: sCache = sCache & "cache": Else: sCache = sCache & "nocache": End If


	sHTML = sHTML & "<!DOCTYPE HTML>"
	sHTML = sHTML & "<html>"
	
	sHTML = sHTML & "<head>"
	sHTML = sHTML & "<meta http-equiv=""X-UA-Compatible"" content=""IE=8 ; IE=9"" />"
	
	sHTML = sHTML & "<meta name=""viewport"" content=""width=device-width,initial-scale=1.0"">"
	
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
	
	htmlHeader2Pinch = sHTML
End Function
%>
