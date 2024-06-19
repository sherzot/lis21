<%
'*******************************************************************************
'概　要：URLEncode文字列をデコード
'引　数：vText		：デコード対象文字列
'　　　：vCharset	：キャラクターセット ["sjis"]["utf8"]["euc"]["jis"]
'出　力：
'戻り値：String		：デコードされた文字列
'備　考：
'更　新：2009/08/06 LIS K.Kokubo 作成
'*******************************************************************************
Function getURLDecode(ByVal vText, ByVal vCharSet)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sSQLErr

	sSQL = "SELECT dbo.func_GetCharactor('" & vText & "','sjis') AS Str;"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sSQLErr)
	If GetRSState(oRS) = True Then
		getURLDecode = oRS.Collect("Str")
	End If
	Call RSClose(oRS)
End Function
%>
