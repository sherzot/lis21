<%
'*******************************************************************************
'概　要：スキル一覧のラヂオボタンを取得
'引　数：vCategoryCode	：カテゴリコード(OS,Application,DevelopmentLanguage,Database)
'　　　：vCode			：チェック中のコード
'　　　：vCols			：一行あたりの列数
'　　　：vName			：inputのname属性値
'戻り値：String
'備　考：
'履　歴：2009/08/05 LIS K.Kokubo 作成
'*******************************************************************************
Function htmlSkillRadio(ByVal vCategoryCode, ByVal vCode, ByVal vCols, ByVal vName)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sSQLErr

	Dim dbCode
	Dim dbSkillName

	Dim sHTML
	Dim aCode
	Dim aFilter
	Dim sChecked
	Dim fWidth
	Dim idx

	sHTML = ""
	fWidth = Round(100 / CInt(vCols), 1)

	aCode = Split(vCode, ",")

	sHTML = sHTML & "<table border=""0"" style=""width:100%;"">" & vbCrLf
	sHTML = sHTML & "<colgroup>"
	For idx = 0 To CInt(vCols) - 1
		sHTML = sHTML & "<col style=""width:" & fWidth & "%;""></col>"
	Next
	sHTML = sHTML & "</colgroup>"
	sHTML = sHTML & "<tbody>"

	sSQL = ""
	sSQL = sSQL & "/* スキル一覧 */" & vbCrLf
	sSQL = sSQL & "SELECT * FROM vw_Skill WHERE CategoryCode = '" & vCategoryCode & "';"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sSQLErr)

	idx = 0
	Do While GetRSState(oRS) = True
		dbCode = oRS.Collect("Code")
		dbSkillName = oRS.Collect("SkillName")

		If idx = 0 Then
			sHTML = sHTML & "<tr>"
		End If

		sChecked = ""
		If UBound(Filter(aCode, dbCode)) >= 0 Then sChecked = " checked"

		sHTML = sHTML & "<td style=""padding:0px;border-width:0px;"">"
		sHTML = sHTML & "<label><input name=""" & vName & """ type=""radio"" value=""" & dbCode & """" & sChecked & ">" & dbSkillName & "</label>"
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

	htmlSkillRadio = sHTML
End Function
%>
