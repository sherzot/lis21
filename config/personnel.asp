<%
response.buffer = "true"

'<�����e�i���X>
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
'</�����e�i���X>

'******************************************************************************
'SQL Server �ݒ� start
'------------------------------------------------------------------------------
'''SQLSERVER2005�e�X�g
'Const DBCNSERVERNAME = "KISUI"		'SQL�T�[�o�[��
'Const DBCNLOGINID    = "Lis21\Administrator"		'SQL���O�C����
'Const DBCNPASSWORD   = "1013Pass2001"	'SQL�p�X���[�h
'Const DBCNDBNAME     = "LISDB"		'SQL�f�[�^�x�[�X��

'''��DB
Const DBCNSERVERNAME = "192.168.151.85"		'SQL�T�[�o�[��
Const DBCNLOGINID    = "shigotonavi"		'SQL���O�C����
Const DBCNPASSWORD   = "1013Pass2000"	'SQL�p�X���[�h
Const DBCNDBNAME     = "LisDB"		'SQL�f�[�^�x�[�X��
'Const DBCNDBNAME     = "LISDB"		'SQL�f�[�^�x�[�X��
'Const DBCNDBNAME     = "TEST_LisDB"		'SQL�f�[�^�x�[�X��
'------------------------------------------------------------------------------
'SQL Server �ݒ� end
'******************************************************************************

'******************************************************************************
'MAIL �ݒ� start
'------------------------------------------------------------------------------
'�V�X�e���Ǘ���
Const MAIL_ADMIN = "kisui@lis21.co.jp"
'���X��\���[��
Const MAIL_LIS = "lis@lis21.co.jp"
'���V�X�e�������[��
Const MAIL_SYSTEM = "kisui@lis21.co.jp"
'���[���T�[�o
Const MAIL_SERVER = "smtp.office365.com"
'Const MAIL_SERVER = "153.153.150.22"
Const Cnt_MailServer = "172.16.1.39"
'------------------------------------------------------------------------------
'MAIL �ݒ� end
'******************************************************************************

'******************************************************************************
'URL �ݒ� start
'------------------------------------------------------------------------------
'�����ƃi�r
Const HTTP_CURRENTURL = "https://www.shigotonavi.co.jp/"
Const HTTPS_CURRENTURL = "https://www.shigotonavi.co.jp/"
Const HTTP_NAVI_CURRENTURL = "http://www.shigotonavi.co.jp/"
Const HTTPS_NAVI_CURRENTURL = "https://www.shigotonavi.co.jp/"
'��������
Const HTTP_RIREKISYO = "http://www.a-rirekisyo.jp"
Const HTTPS_RIREKISYO = "https://www.a-rirekisyo.jp"
'���X�g�o
Const HTTP_LIS_CURRENTURL = "http://www.lis21.co.jp/"
Const HTTPS_LIS_CURRENTURL = "https://www.lis21.co.jp/"
'�l�ލ̗p
Const HTTP_JINZAI_CURRENTURL = "http://jinzai.shigotonavi.co.jp/"
'�Г��V�X�e��
Const HTTP_BI_CURRENTURL = "http://bi.lis21.co.jp/"
'�����ƃi�r���o�C��
Const HTTP_NAVI_MOBILE = "http://m.shigotonavi.jp/"
Const HTTPS_NAVI_MOBILE = "https://m.shigotonavi.jp/"
'�����ƃi�r�X�}�z
Const HTTP_SP = "http://sp.shigotonavi.jp/"
Const HTTPS_SP = "https://sp.shigotonavi.jp/"
'�В��i�r
Const HTTP_EX = "http://www.shigotonavi.co.jp/ex/"
Const HTTPS_EX = "https://www.shigotonavi.co.jp/ex/"
'�����ƃi�rFacebook�y�[�W
Const HTTP_FB = "http://www.facebook.com/shigoto"

'�N�x���Ƃɕς��V���̗p�y�[�W�̂t�q�k
Dim HTTP_SHINSOTSU: HTTP_SHINSOTSU = HTTP_CURRENTURL & "lis/recruit_shinsotsu08_index.asp"	'�V���s�n�o

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
'URL �ݒ� end
'******************************************************************************

'******************************************************************************
'�O���[�o���ϐ� start
'------------------------------------------------------------------------------
'���O�C�����̑㗝�X�R�[�h
Dim G_AGCCODE			:G_AGCCODE = Session("agencycode")
'���O�C�����̑㗝�X���_�ԍ�
Dim G_AGCBRANCH			:G_AGCBRANCH = Session("agencybranch")
'���O�C�����̃��[�U�h�c
Dim G_USERID			:G_USERID = Session("userid")
'���O�C�����̃��[�U���
Dim G_USERTYPE			:G_USERTYPE = Session("usertype")
'���O�C������Ƃ̗^�M�t���O
Dim G_YOSHIN            :G_YOSHIN = Session("YoshinFlag")
'���O�C������Ƃ̊�Ƌ敪
Dim G_COMPANYKBN		:G_COMPANYKBN = Session("companykbn")
'���O�C������Ƃ̃��C�Z���X���
Dim G_PLANTYPE			:G_PLANTYPE = Session("plantype")
'���O�C������Ƃ̃��C�Z���X�\�����݃R�[�h
Dim G_APPLICATIONCODE	:G_APPLICATIONCODE = Session("applicationcode")
'���O�C������Ƃ̋����C�Z���X�\�����݃R�[�h
Dim G_OLDAPPLICATIONCODE:G_OLDAPPLICATIONCODE = Session("oldapplicationcode")
'���O�C������Ƃ̋����C�Z���X���
Dim G_OLDPLANTYPE		:G_OLDPLANTYPE = Session("oldplantype")
'���O�C������Ƃ̃��C�Z���X�̗L���t���O
Dim G_USEFLAG			:G_USEFLAG = Session("useflag")
'���O�C������Ƃ̃��C�Z���X�̋��l�[�f�ڗL���t���O
Dim G_PUBLICFLAG		:G_PUBLICFLAG = Session("publicflag")
'���O�C������Ƃ̃��C�Z���X���؂�Ă��Ă����[���\�t���O
Dim G_MAILREADFLAG		:G_MAILREADFLAG = Session("mailreadflag")
'���O�C������Ƃ̌f�ډ\���l�[�ʐ^��
Dim G_IMAGELIMIT		:G_IMAGELIMIT = Session("imagelimit")
'���O�C������Ƃ̋����C�Z���X�̌f�ډ\���l�[�ʐ^��
Dim G_OLDIMAGELIMIT		:G_OLDIMAGELIMIT = Session("oldimagelimit")
'���O�C������Ƃ̃C���^�r���[�f�ډۃt���O
Dim G_INTERVIEWFLAG		:G_INTERVIEWFLAG = Session("interviewflag")
'���O�C������Ƃ̋����C�Z���X�̃C���^�r���[�f�ډۃt���O
Dim G_OLDINTERVIEWFLAG	:G_OLDINTERVIEWFLAG = Session("oldinterviewflag")
'���O�C������Ƃ̔h���F�t���O
Dim G_TEMPPERMITFLAG	:G_TEMPPERMITFLAG = Session("temppermitflag")
'���O�C������Ƃ̏Љ�F�t���O
Dim G_INTROPERMITFLAG	:G_INTROPERMITFLAG = Session("intropermitflag")
'���l�[�ڍ׌����p�p�����[�^
Dim G_PARAMSEARCHORDER	:G_PARAMSEARCHORDER = Session("paramsearchorder")
'�v�d�a�T�[�o��
Dim G_WEBSERVERNAME		:G_WEBSERVERNAME = Request.ServerVariables("SERVER_NAME")
'�p�����[�^
Dim G_QUERYSTRING		:G_QUERYSTRING = Request.ServerVariables("QUERY_STRING")
'���݂̊��S�t�q�k
Dim G_URL
G_URL = "http://" & G_WEBSERVERNAME & Request.ServerVariables("URL")
'���݂̊��S�t�q�k(�r�r�k)
Dim G_URLS
G_URLS = "https://" & G_WEBSERVERNAME & Request.ServerVariables("URL")
If G_QUERYSTRING <> "" Then G_URL = G_URL & "?" & G_QUERYSTRING
'���t�@���[
Dim G_REFERER			:G_REFERER = Request.ServerVariables("HTTP_REFERER")
'�h�o�A�h���X
Dim G_IPADDRESS			:G_IPADDRESS = Request.ServerVariables("REMOTE_ADDR")
'���[�U�[�G�[�W�F���g
Dim G_USERAGENT			:G_USERAGENT = Request.ServerVariables("HTTP_USER_AGENT")
'������������
Dim G_FLGRESUME			:G_FLGRESUME = False
If InStr(Request.ServerVariables("URL"), "www.a-rirekisyo.jp") <> 0 Then G_FLGRESUME = True
If InStr(Request.ServerVariables("URL"), "/resume/") <> 0 Then G_FLGRESUME = True
'�r�r�k�t���O
Dim G_SSLFLAG
If Request.ServerVariables("HTTPS") = "on" Then
	G_SSLFLAG = True
Else
	G_SSLFLAG = False
End If

'�ŏ��̖K��̂��������i�L���Ȃǁj
Dim G_ADVERTISEMENT
'1.Cookie������Ύ擾
If Session("advertisement") = "" Then
	Session("advertisement") = GetCookie("advertisement")
End If
'2.�L���p�����[�^������Ύ擾
If Session("advertisement") = "" And (InStr(G_QUERYSTRING, "rf=") <> 0) Then
	If WriteCookie("advertisement", G_URL) = True Then
		Response.Cookies("advertisement") = G_URL
		Response.Cookies("advertisement").Expires = Date + 30
	End If
	Session("advertisement") = G_URL
End If
'3.���t�@���[���i�r�T�C�g�ȊO�̂��̂ł���Ύ擾
If Session("advertisement") = "" And InStr(G_REFERER, G_WEBSERVERNAME) = 0 Then
	If WriteCookie("advertisement", G_REFERER) = True Then
		Response.Cookies("advertisement") = G_REFERER
		Response.Cookies("advertisement").Expires = Date + 30
	End If
	Session("advertisement") = G_REFERER
End If
G_ADVERTISEMENT = Session("advertisement")


'******************************************************************************
'����ƌ����L���W�v�p�̕ϐ����`
'�P�T�Ԉȓ��̃A�N�Z�X�Œ��߂ǂ̃����f�B���O�y�[�W�փA�N�Z�X���ǂ�������
'******************************************************************************
Dim G_ADVISITERURL
'1.Cookie������Ύ擾
If Session("advisiterurl") = "" Then
	Session("advisiterurl") = GetCookie("advisiterurl")
End If
'2.�L���p�����[�^������Ύ擾
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
'�O���[�o���ϐ� end
'******************************************************************************

'******************************************************************************
'�Œ�l start
'------------------------------------------------------------------------------
'Google�e�X�g
'Const GOOGLEMAPSAPIKEY = "ABQIAAAAGfCbzPsVkm3lk14QtpM60RQeCxK22fcwiw3345Yi-qh3jiDOqRRsOexODagFuOmqemdBR0_jSZBQAA"
'Google�{��
Const GOOGLEMAPSAPIKEY = "ABQIAAAAGfCbzPsVkm3lk14QtpM60RQTa3R4hR7qmMa_Tvsti0VigvA1zhSlA6kXtvuSCAVH21Scg8440HDekA"

'������FDF�e���v���[�g�p�X
Const RESUME_VER3 = "F:\asp-source\�����ƃi�r\staff\resume_ver3.fdf"
Const CAREERSHEET_OFFICE = "F:\asp-source\�����ƃi�r\staff\careersheet_office.fdf"
Const CAREERSHEET_IT = "F:\asp-source\�����ƃi�r\staff\careersheet-it.fdf"

'�^�C�g���ɕt�^����Œ蕶����
Const TITLE_STR = "&nbsp;�y�]�E�E���l�T�C�g�����ƃi�r�z"
Const TITLE_CMP = "&nbsp;�y���l�L�������ƃi�r�z"

'�R���v���p�o�c�e�t�@�C���ۑ���
Dim ConpriFolder	:	ConpriFolder = "F:\asp-source\�����ƃi�r\Conpri\infile\"

Dim ReportFolder	:	ReportFolder = "F:\asp-source\���[�t�H�[�}�b�g"

'���l�[�摜�ő吔
Const MAXORDERIMG = 100
'------------------------------------------------------------------------------
'�Œ�l end
'******************************************************************************

'�G���[�������擾�p
Session("errorpagereferer") = Request.ServerVariables("HTTP_REFERER")
Session("errorpage") = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")

'�o�i�[�R���g���[���p�ϐ� ad_banner_control/ad_banner.asp
Dim gBannerSQL
Dim gBannerRS
Dim gBannerCode
Dim gBannerFileName
Dim gBannerURL

'�A�t�B���G�C�g����̃A�N�Z�X
If Request.QueryString("bt") = "af" Then Session("flgaffiliate") = "1"

'<Cookie����G���[���p�֐�>
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
'</Cookie����G���[���p�֐�>
%>