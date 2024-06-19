<%
'*******************************************************************************
'概　要：業種一覧のチェックボックスを取得
'引　数：vCode	：チェック中のコード
'　　　：vCols	：一行あたりの列数
'　　　：vName	：inputのname属性値
'戻り値：String
'備　考：
'履　歴：2009/08/05 LIS K.Kokubo 作成
'*******************************************************************************
Function htmlIndustryTypeRadio(ByVal vCode, ByVal vCols, ByVal vName)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sSQLErr

	Dim dbIndustryTypeCode
	Dim dbIndustryTypeName

	Dim sHTML
	Dim aCode
	Dim aFilter
	Dim sChecked
	Dim fWidth
	Dim idx

	sHTML = ""
	fWidth = Round(100 / CInt(vCols), 1)

	sSQL = ""
	sSQL = sSQL & "/* 業種一覧 */" & vbCrLf
	sSQL = sSQL & "EXEC sp_GetList 'IndustryType';"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sSQLErr)

	aCode = Split(vCode, ",")

	sHTML = sHTML & "<table border=""0"" style=""width:100%;"">" & vbCrLf
	sHTML = sHTML & "<colgroup>"
	For idx = 0 To CInt(vCols) - 1
		sHTML = sHTML & "<col style=""width:" & fWidth & "%;""></col>"
	Next
	sHTML = sHTML & "</colgroup>"
	sHTML = sHTML & "<tbody>"

	idx = 0
	Do While GetRSState(oRS) = True
		dbIndustryTypeCode = oRS.Collect("Code")
		dbIndustryTypeName = oRS.Collect("Detail")

		If Right(dbIndustryTypeCode, 2) = "00" Then
			If idx <> 0 Then sHTML = sHTML & "</tr>"
			sHTML = sHTML & "<tr>"
			sHTML = sHTML & "<td colspan=""" & vCols & """ style=""padding:4px;border-width:0px;font-weight:bold;"">" & dbIndustryTypeName & "</td>"
			sHTML = sHTML & "</tr>"
			idx = 0

			oRS.MoveNext
			If GetRSState(oRS) = True Then
				dbIndustryTypeCode = oRS.Collect("Code")
				dbIndustryTypeName = oRS.Collect("Detail")
			Else
				Exit Do
			End If
		End If

		If idx = 0 Then
			sHTML = sHTML & "<tr>"
		End If

		sChecked = ""
		If UBound(Filter(aCode, dbIndustryTypeCode)) >= 0 Then sChecked = " checked"

		sHTML = sHTML & "<td style=""padding:0px;border-width:0px;"">"
		sHTML = sHTML & "<label><input name=""" & vName & """ type=""radio"" value=""" & dbIndustryTypeCode & """" & sChecked & ">" & dbIndustryTypeName & "</label>"
		sHTML = sHTML & "</td>"

		oRS.MoveNext

		If GetRSState(oRS) = True Then
			If oRS.Collect("Code") = "Z99" Then oRS.MoveNext
		End If

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

	htmlIndustryTypeRadio = sHTML
End Function
%>
