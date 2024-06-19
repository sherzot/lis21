<%
'*******************************************************************************
'概　要：オーバーチュアタグ取得
'引　数：
'出　力：
'戻り値：String
'備　考：
'履　歴：2010/05/11 LIS K.Kokubo 作成
'*******************************************************************************
Function scrOverT_opti_shopper()
	Dim sScript

'	sScript = vbCrLf
'	sScript = sScript & "<script type=""text/javascript"" language=""javascript"">" & vbCrLf
'	sScript = sScript & "<!-- Overture Services, Inc" & vbCrLf
'sScript = sScript & "/*" & vbCrLf
'	sScript = sScript & "var pm_tagname = 'shopperTag.txt';" & vbCrLf
'	sScript = sScript & "var pm_tagversion = '1.3';" & vbCrLf
'	sScript = sScript & "window.pm_customData = new Object();" & vbCrLf
'	sScript = sScript & "window.pm_customData.segment='name=shopper, transId=';" & vbCrLf
'sScript = sScript & "*/" & vbCrLf
'	sScript = sScript & "// -->" & vbCrLf
'	sScript = sScript & "</script>" & vbCrLf

	scrOverT_opti_shopper = sScript
End Function
%>
