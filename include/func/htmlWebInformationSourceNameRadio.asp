<%
'*******************************************************************************
'概　要：同業他社ＷＥサイト一覧のラヂオボタンを取得
'引　数：vSourceName：チェック中のコード
'　　　：vCols		：一行あたりの列数
'　　　：vName		：inputのname属性値
'　　　：vAttribute	：inputの追加属性
'戻り値：String
'備　考：
'履　歴：2009/09/09 LIS K.Kokubo 作成
'*******************************************************************************
Function htmlWebInformationSourceNameRadio(ByVal vSourceName, ByVal vCols, ByVal vName, ByVal vAttribute)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sSQLErr

	Dim dbWebInformationSourceTypeCode
	Dim dbWebInformationSourceTypeName

	Dim sHTML
	Dim aName
	Dim aFilter
	Dim sChecked
	Dim fWidth
	Dim idx

	sHTML = ""
	fWidth = Round(100 / CInt(vCols), 1)

	If vAttribute <> "" Then vAttribute = " " & vAttribute

	sSQL = ""
	sSQL = sSQL & "/* 同業他社ＷＥサイト一覧 */" & vbCrLf
	sSQL = sSQL & "SELECT * FROM vw_WebInformationSourceType ORDER BY WebInformationSourceTypeCode;"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sSQLErr)

	aName = Split(vSourceName, ",")

	sHTML = sHTML & "<table border=""0"" style=""width:100%;"">" & vbCrLf
	sHTML = sHTML & "<colgroup>"
	For idx = 0 To CInt(vCols) - 1
		sHTML = sHTML & "<col style=""width:" & fWidth & "%;""></col>"
	Next
	sHTML = sHTML & "</colgroup>"
	sHTML = sHTML & "<tbody>"

	idx = 0
	Do While GetRSState(oRS) = True
		dbWebInformationSourceTypeCode = oRS.Collect("WebInformationSourceTypeCode")
		dbWebInformationSourceTypeName = oRS.Collect("WebInformationSourceTypeName")

		If idx = 0 Then
			sHTML = sHTML & "<tr>"
		End If

		sChecked = ""
		If UBound(Filter(aName, dbWebInformationSourceTypeName)) >= 0 Then sChecked = " checked"

		sHTML = sHTML & "<td style=""padding:0px;border-width:0px;"">"
		sHTML = sHTML & "<label><input name=""" & vName & """ type=""radio"" value=""" & dbWebInformationSourceTypeName & """" & sChecked & "" & vAttribute & ">" & dbWebInformationSourceTypeName & "</label>"
		sHTML = sHTML & "</td>"

		oRS.MoveNext

		If GetRSState(oRS) = False Or idx = CInt(vCols) - 1 Then
			sHTML = sHTML & "</tr>"
			idx = 0
		Else
			idx = idx + 1
		End If
	Loop
	Call RSClose(oRS)

	sHTML = sHTML & "</tbody>"
	sHTML = sHTML & "</table>"

	htmlWebInformationSourceNameRadio = sHTML
End Function
%>
