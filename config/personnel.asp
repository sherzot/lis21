<%
response.buffer = "true"

'<メンテナンス>
'If Request.ServerVariables("URL") <> "/maintenance/index.asp" _
'And (Now >= "2015/04/19 09:00:00" And Now <= "2015/04/19 23:00:00") Then
	'Response.Status="302 Found"
'	If Now <= "2015/04/19 23:00:00" Then
	'	Response.AddHeader "Retry-After","Sun, 19 Apr 2015 23:00:00 GMT"
'	End If
'If InStr(Request.ServerVariables("REMOTE_HOST"),"192.168.") = 0 Then
	'Response.AddHeader "Location", "http://www.shigotonavi.co.jp/maintenance/"
'END IF
'End If
'</メンテナンス>

'******************************************************************************
'SQL Server 設定 start
'------------------------------------------------------------------------------
'''SQLSERVER2005テスト
'Const DBCNSERVERNAME = "KISUI"		'SQLサーバー名
'Const DBCNLOGINID    = "Lis21\Administrator"		'SQLログイン名
'Const DBCNPASSWORD   = "1013Pass2001"	'SQLパスワード
'Const DBCNDBNAME     = "LISDB"		'SQLデータベース名

'''現DB
Const DBCNSERVERNAME = "192.168.151.85"		'SQLサーバー名
Const DBCNLOGINID    = "shigotonavi"		'SQLログイン名
Const DBCNPASSWORD   = "1013Pass2000"	'SQLパスワード
Const DBCNDBNAME     = "LisDB"		'SQLデータベース名
'Const DBCNDBNAME     = "LISDB"		'SQLデータベース名
'Const DBCNDBNAME     = "TEST_LisDB"		'SQLデータベース名
'------------------------------------------------------------------------------
'SQL Server 設定 end
'******************************************************************************

'******************************************************************************
'MAIL 設定 start
'------------------------------------------------------------------------------
'システム管理者
Const MAIL_ADMIN = "kisui@lis21.co.jp"
'リス代表メール
Const MAIL_LIS = "lis@lis21.co.jp"
'企画システム室メール
Const MAIL_SYSTEM = "kisui@lis21.co.jp"
'メールサーバ
Const MAIL_SERVER = "smtp.office365.com"
'Const MAIL_SERVER = "153.153.150.22"
Const Cnt_MailServer = "172.16.1.39"
'------------------------------------------------------------------------------
'MAIL 設定 end
'******************************************************************************

'******************************************************************************
'URL 設定 start
'------------------------------------------------------------------------------
'しごとナビ
Const HTTP_CURRENTURL = "https://www.shigotonavi.co.jp/"
Const HTTPS_CURRENTURL = "https://www.shigotonavi.co.jp/"
Const HTTP_NAVI_CURRENTURL = "http://www.shigotonavi.co.jp/"
Const HTTPS_NAVI_CURRENTURL = "https://www.shigotonavi.co.jp/"
'＠履歴書
Const HTTP_RIREKISYO = "http://www.a-rirekisyo.jp"
Const HTTPS_RIREKISYO = "https://www.a-rirekisyo.jp"
'リスＨＰ
Const HTTP_LIS_CURRENTURL = "http://www.lis21.co.jp/"
Const HTTPS_LIS_CURRENTURL = "https://www.lis21.co.jp/"
'人材採用
Const HTTP_JINZAI_CURRENTURL = "http://jinzai.shigotonavi.co.jp/"
'社内システム
Const HTTP_BI_CURRENTURL = "http://bi.lis21.co.jp/"
'しごとナビモバイル
Const HTTP_NAVI_MOBILE = "http://m.shigotonavi.jp/"
Const HTTPS_NAVI_MOBILE = "https://m.shigotonavi.jp/"
'しごとナビスマホ
Const HTTP_SP = "http://sp.shigotonavi.jp/"
Const HTTPS_SP = "https://sp.shigotonavi.jp/"
'社長ナビ
Const HTTP_EX = "http://www.shigotonavi.co.jp/ex/"
Const HTTPS_EX = "https://www.shigotonavi.co.jp/ex/"
'しごとナビFacebookページ
Const HTTP_FB = "http://www.facebook.com/shigoto"

'年度ごとに変わる新卒採用ページのＵＲＬ
Dim HTTP_SHINSOTSU: HTTP_SHINSOTSU = HTTP_CURRENTURL & "lis/recruit_shinsotsu08_index.asp"	'新卒ＴＯＰ

Dim BASEURL
Dim NAVI_BASEURL
Dim CURRENTURL

If Request.ServerVariables("HTTPS") = "on" Then
	BASEURL = HTTPS_CURRENTURL
	NAVI_BASEURL = HTTPS_CURRENTURL
Else
	BASEURL = HTTP_CURRENTURL
	NAVI_BASEURL = HTTP_CURRENTURL
End If

CURRENTURL = Request.ServerVariables("URL")
'------------------------------------------------------------------------------
'URL 設定 end
'******************************************************************************

'******************************************************************************
'グローバル変数 start
'------------------------------------------------------------------------------
'ログイン中の代理店コード
Dim G_AGCCODE			:G_AGCCODE = Session("agencycode")
'ログイン中の代理店拠点番号
Dim G_AGCBRANCH			:G_AGCBRANCH = Session("agencybranch")
'ログイン中のユーザＩＤ
Dim G_USERID			:G_USERID = Session("userid")
'ログイン中のユーザ種類
Dim G_USERTYPE			:G_USERTYPE = Session("usertype")
'ログイン中企業の与信フラグ
Dim G_YOSHIN            :G_YOSHIN = Session("YoshinFlag")
'ログイン中企業の企業区分
Dim G_COMPANYKBN		:G_COMPANYKBN = Session("companykbn")
'ログイン中企業のライセンス種類
Dim G_PLANTYPE			:G_PLANTYPE = Session("plantype")
'ログイン中企業のライセンス申し込みコード
Dim G_APPLICATIONCODE	:G_APPLICATIONCODE = Session("applicationcode")
'ログイン中企業の旧ライセンス申し込みコード
Dim G_OLDAPPLICATIONCODE:G_OLDAPPLICATIONCODE = Session("oldapplicationcode")
'ログイン中企業の旧ライセンス種類
Dim G_OLDPLANTYPE		:G_OLDPLANTYPE = Session("oldplantype")
'ログイン中企業のライセンスの有効フラグ
Dim G_USEFLAG			:G_USEFLAG = Session("useflag")
'ログイン中企業のライセンスの求人票掲載有効フラグ
Dim G_PUBLICFLAG		:G_PUBLICFLAG = Session("publicflag")
'ログイン中企業のライセンスが切れていてもメール可能フラグ
Dim G_MAILREADFLAG		:G_MAILREADFLAG = Session("mailreadflag")
'ログイン中企業の掲載可能求人票写真数
Dim G_IMAGELIMIT		:G_IMAGELIMIT = Session("imagelimit")
'ログイン中企業の旧ライセンスの掲載可能求人票写真数
Dim G_OLDIMAGELIMIT		:G_OLDIMAGELIMIT = Session("oldimagelimit")
'ログイン中企業のインタビュー掲載可否フラグ
Dim G_INTERVIEWFLAG		:G_INTERVIEWFLAG = Session("interviewflag")
'ログイン中企業の旧ライセンスのインタビュー掲載可否フラグ
Dim G_OLDINTERVIEWFLAG	:G_OLDINTERVIEWFLAG = Session("oldinterviewflag")
'ログイン中企業の派遣認可フラグ
Dim G_TEMPPERMITFLAG	:G_TEMPPERMITFLAG = Session("temppermitflag")
'ログイン中企業の紹介認可フラグ
Dim G_INTROPERMITFLAG	:G_INTROPERMITFLAG = Session("intropermitflag")
'求人票詳細検索用パラメータ
Dim G_PARAMSEARCHORDER	:G_PARAMSEARCHORDER = Session("paramsearchorder")
'ＷＥＢサーバ名
Dim G_WEBSERVERNAME		:G_WEBSERVERNAME = Request.ServerVariables("SERVER_NAME")
'パラメータ
Dim G_QUERYSTRING		:G_QUERYSTRING = Request.ServerVariables("QUERY_STRING")
'現在の完全ＵＲＬ
Dim G_URL
G_URL = "http://" & G_WEBSERVERNAME & Request.ServerVariables("URL")
'現在の完全ＵＲＬ(ＳＳＬ)
Dim G_URLS
G_URLS = "https://" & G_WEBSERVERNAME & Request.ServerVariables("URL")
If G_QUERYSTRING <> "" Then G_URL = G_URL & "?" & G_QUERYSTRING
'リファラー
Dim G_REFERER			:G_REFERER = Request.ServerVariables("HTTP_REFERER")
'ＩＰアドレス
Dim G_IPADDRESS			:G_IPADDRESS = Request.ServerVariables("REMOTE_ADDR")
'ユーザーエージェント
Dim G_USERAGENT			:G_USERAGENT = Request.ServerVariables("HTTP_USER_AGENT")
'＠履歴書判別
Dim G_FLGRESUME			:G_FLGRESUME = False
If InStr(Request.ServerVariables("URL"), "www.a-rirekisyo.jp") <> 0 Then G_FLGRESUME = True
If InStr(Request.ServerVariables("URL"), "/resume/") <> 0 Then G_FLGRESUME = True
'ＳＳＬフラグ
Dim G_SSLFLAG
If Request.ServerVariables("HTTPS") = "on" Then
	G_SSLFLAG = True
Else
	G_SSLFLAG = False
End If

'最初の訪問のきっかけ（広告など）
Dim G_ADVERTISEMENT
'1.Cookieがあれば取得
If Session("advertisement") = "" Then
	Session("advertisement") = GetCookie("advertisement")
End If
'2.広告パラメータがあれば取得
If Session("advertisement") = "" And (InStr(G_QUERYSTRING, "rf=") <> 0) Then
	If WriteCookie("advertisement", G_URL) = True Then
		Response.Cookies("advertisement") = G_URL
		Response.Cookies("advertisement").Expires = Date + 30
	End If
	Session("advertisement") = G_URL
End If
'3.リファラーがナビサイト以外のものであれば取得
If Session("advertisement") = "" And InStr(G_REFERER, G_WEBSERVERNAME) = 0 Then
	If WriteCookie("advertisement", G_REFERER) = True Then
		Response.Cookies("advertisement") = G_REFERER
		Response.Cookies("advertisement").Expires = Date + 30
	End If
	Session("advertisement") = G_REFERER
End If
G_ADVERTISEMENT = Session("advertisement")


'******************************************************************************
'■企業向け広告集計用の変数を定義
'１週間以内のアクセスで直近どのランディングページへアクセスかどうか判別
'******************************************************************************
Dim G_ADVISITERURL
'1.Cookieがあれば取得
If Session("advisiterurl") = "" Then
	Session("advisiterurl") = GetCookie("advisiterurl")
End If
'2.広告パラメータがあれば取得
If Session("advisiterurl") = "" And InStr(G_QUERYSTRING, "rf=") > 0 and InStr(G_REFERER, G_WEBSERVERNAME) = 0 Then
	If WriteCookie("advisiterurl", G_URL) = True Then
		Response.Cookies("advisiterurl") = G_URL
		Response.Cookies("advisiterurl").Expires = Date + 7
	End If
	Session("advisiterurl") = G_URL
elseif Session("advisiterurl") <> "" and Len(G_REFERER) > 0 and InStr(G_REFERER, G_WEBSERVERNAME) = 0 and InStr(G_QUERYSTRING, "rf=") > 0 Then
	If WriteCookie("advisiterurl", G_URL) = True Then
		Response.Cookies("advisiterurl") = G_URL
		Response.Cookies("advisiterurl").Expires = Date + 7
	End If
	Session("advisiterurl") = G_URL
End If
G_ADVISITERURL = Session("advisiterurl")
'******************************************************************************

'------------------------------------------------------------------------------
'グローバル変数 end
'******************************************************************************

'******************************************************************************
'固定値 start
'------------------------------------------------------------------------------
'Googleテスト
'Const GOOGLEMAPSAPIKEY = "ABQIAAAAGfCbzPsVkm3lk14QtpM60RQeCxK22fcwiw3345Yi-qh3jiDOqRRsOexODagFuOmqemdBR0_jSZBQAA"
'Google本番
Const GOOGLEMAPSAPIKEY = "ABQIAAAAGfCbzPsVkm3lk14QtpM60RQTa3R4hR7qmMa_Tvsti0VigvA1zhSlA6kXtvuSCAVH21Scg8440HDekA"

'履歴書FDFテンプレートパス
Const RESUME_VER3 = "F:\asp-source\しごとナビ\staff\resume_ver3.fdf"
Const CAREERSHEET_OFFICE = "F:\asp-source\しごとナビ\staff\careersheet_office.fdf"
Const CAREERSHEET_IT = "F:\asp-source\しごとナビ\staff\careersheet-it.fdf"

'タイトルに付与する固定文字列
Const TITLE_STR = "&nbsp;【転職・求人サイトしごとナビ】"
Const TITLE_CMP = "&nbsp;【求人広告しごとナビ】"

'コンプリ用ＰＤＦファイル保存先
Dim ConpriFolder	:	ConpriFolder = "F:\asp-source\しごとナビ\Conpri\infile\"

Dim ReportFolder	:	ReportFolder = "F:\asp-source\帳票フォーマット"

'求人票画像最大数
Const MAXORDERIMG = 100
'------------------------------------------------------------------------------
'固定値 end
'******************************************************************************

'エラー発生源取得用
Session("errorpagereferer") = Request.ServerVariables("HTTP_REFERER")
Session("errorpage") = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")

'バナーコントロール用変数 ad_banner_control/ad_banner.asp
Dim gBannerSQL
Dim gBannerRS
Dim gBannerCode
Dim gBannerFileName
Dim gBannerURL

'アフィリエイトからのアクセス
If Request.QueryString("bt") = "af" Then Session("flgaffiliate") = "1"

'<Cookie操作エラー回避用関数>
Function GetCookie(ByVal vKey)
	On Error Resume Next
	GetCookie = Request.Cookies(vKey)
End Function

Function WriteCookie(ByVal vKey, ByVal vData)
	On Error Resume Next
	WriteCookie = True
	Response.Cookies(vKey) = vData
	If Err.Number <> 0 Then WriteCookie = False
End Function
'</Cookie操作エラー回避用関数>
%>
