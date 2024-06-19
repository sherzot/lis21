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
'*******************************************************************************
Function htmlHeader(ByVal vSite, ByVal vTitle, ByVal vKeywords, ByVal vDescription, ByVal vAddHead, ByVal vIndexFlag, ByVal vFollowFlag, ByVal vArchiveFlag, ByVal vCacheFlag, ByVal vBodyAttribute)
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

'	Response.Charset = "UTF-8"
	'IEはapplication/xhtml+xmlのmimeを認識できない
	'If InStr(G_USERAGENT, "MSIE") > 0 Then
	'	'IE
	'	sContentType = "text/html"
	'Else
	'	sContentType = "application/xhtml+xml"
	'End If
'	sContentType = "text/html"
'	Response.ContentType = sContentType

	sHTML = sHTML & "<!DOCTYPE HTML>"
	sHTML = sHTML & "<html>"
	sHTML = sHTML & "<head>"
	'If vSite <> "" Then sHTML = sHTML & "<base href=""" & vSite & """ />"
	sHTML = sHTML & "<meta http-equiv=""content-type"" content=""text/html; charset=shift_jis"">"
	sHTML = sHTML & "<meta http-equiv=""content-script-type"" content=""text/javascript"">"
	sHTML = sHTML & "<meta http-equiv=""content-style-type"" content=""text/css"">"
	sHTML = sHTML & "<meta name=""robots"" content=""" & sRobots & """>"
	sHTML = sHTML & "<meta name=""googlebot"" content=""" & sRobots & """>"
	sHTML = sHTML & "<meta name=""keywords"" content=""" & vKeywords & """>"
	sHTML = sHTML & "<meta name=""description"" content=""" & vDescription & """>"
	'sHTML = sHTML & "<meta http-equiv=""X-UA-Compatible"" content=""IE=EmulateIE7"">"
	sHTML = sHTML & vAddHead
	sHTML = sHTML & "<title>" & vTitle & "</title>"
%>
<!--[if lt IE 9]>
<script src="/script/html5.js"></script>
<![endif]-->
<%

	sHTML = sHTML & "</head>"
	If vBodyAttribute <> "" Then vBodyAttribute = " " & vBodyAttribute
	sHTML = sHTML & "<body" & vBodyAttribute & ">" & vbCrLf

	htmlHeader = sHTML
End Function
%>
