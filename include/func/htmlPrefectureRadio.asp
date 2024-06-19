<%
'*******************************************************************************
'概　要：都道府県一覧のチェックボックスを取得
'引　数：vCode			：チェック中のコード
'　　　：vCols			：一行あたりの列数
'　　　：vName			：inputのname属性値
'　　　：vAttribute		：inputの追加属性
'　　　：vForeignFlag	：海外表示フラグ
'戻り値：String
'備　考：
'履　歴：2009/08/05 LIS K.Kokubo 作成
'　　　：2009/08/21 LIS K.Kokubo vForeignFlag追加
'*******************************************************************************
Function htmlPrefectureRadio(ByVal vCode, ByVal vCols, ByVal vName, ByVal vAttribute, ByVal vForeignFlag)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sSQLErr

	Dim dbPrefectureCode
	Dim dbPrefectureName

	Dim sHTML
	Dim aCode
	Dim aFilter
	Dim sChecked
	Dim fWidth
	Dim idx

	sHTML = ""
	fWidth = Round(100 / CInt(vCols), 1)

	If vAttribute <> "" Then vAttribute = " " & vAttribute

	sSQL = ""
	If vForeignFlag = "1" Then
		sSQL = sSQL & "/* 都道府県(海外含む)一覧 */" & vbCrLf
		sSQL = sSQL & "EXEC up_LstPrefectureAll;"
	Else
		sSQL = sSQL & "/* 都道府県一覧 */" & vbCrLf
		sSQL = sSQL & "EXEC up_LstPrefecture;"
	End If
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
		dbPrefectureCode = oRS.Collect("PrefectureCode")
		dbPrefectureName = oRS.Collect("PrefectureName")

		If idx = 0 Then
			sHTML = sHTML & "<tr>"
		End If

		sChecked = ""
		If UBound(Filter(aCode, dbPrefectureCode)) >= 0 Then sChecked = " checked"

		sHTML = sHTML & "<td style=""padding:0px;border-width:0px;"">"
		sHTML = sHTML & "<label><input name=""" & vName & """ type=""radio"" value=""" & dbPrefectureCode & """" & sChecked & "" & vAttribute & ">" & dbPrefectureName & "</label>"
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

	htmlPrefectureRadio = sHTML
End Function
%>
