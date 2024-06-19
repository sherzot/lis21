<%
'*******************************************************************************
'概　要：勤務形態一覧のラヂオボタンを取得(新卒を除く)
'引　数：vCode			：チェック中のコード
'　　　：vCols			：一行あたりの列数
'　　　：vName			：inputのname属性値
'戻り値：String
'備　考：
'履　歴：2011/01/25 LIS K.Kokubo 作成
'*******************************************************************************
Function htmlWorkingTypeRadio2(ByVal vCode, ByVal vCols, ByVal vName)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sSQLErr

	Dim dbWorkingTypeCode
	Dim dbWorkingTypeName

	Dim sHTML
	Dim aCode
	Dim aFilter
	Dim sChecked
	Dim fWidth
	Dim idx

	sHTML = ""
	If vCols > 0 Then fWidth = Round(100 / CInt(vCols), 1)

	sSQL = ""
	sSQL = sSQL & "/* 勤務形態一覧 */" & vbCrLf
	sSQL = sSQL & "EXEC sp_GetList 'WorkingType';"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sSQLErr)
	If GetRSState(oRS) = True Then
		oRS.Filter = "Code <> '100'"
	End If

	aCode = Split(vCode, ",")

	sHTML = sHTML & "<table border=""0"" style=""width:100%;"">" & vbCrLf
	If CInt(vCols) > 0 Then
		sHTML = sHTML & "<colgroup>"
		For idx = 0 To CInt(vCols) - 1
			sHTML = sHTML & "<col style=""width:" & fWidth & "%;""></col>"
		Next
		sHTML = sHTML & "</colgroup>"
	End If
	sHTML = sHTML & "<tbody>"

	idx = 0
	Do While GetRSState(oRS) = True
		dbWorkingTypeCode = oRS.Collect("Code")
		dbWorkingTypeName = oRS.Collect("Detail")
		If dbWorkingTypeCode = "005" Then
			dbWorkingTypeName = "パート,アルバイト"
		ElseIf dbWorkingTypeCode = "006" Then
			dbWorkingTypeName = "ＳＯＨＯ"
		End If

		If idx = 0 Then
			sHTML = sHTML & "<tr>"
		End If

		sChecked = ""
		If UBound(Filter(aCode, dbWorkingTypeCode)) >= 0 Then sChecked = " checked"

		sHTML = sHTML & "<td style=""padding:0px;border-width:0px;"">"
		sHTML = sHTML & "<label><input name=""" & vName & """ type=""radio"" value=""" & dbWorkingTypeCode & """" & sChecked & ">" & dbWorkingTypeName & "</label>"
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

	htmlWorkingTypeRadio2 = sHTML
End Function
%>
