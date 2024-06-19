<%
'*******************************************************************************
'概　要：都道府県一覧の<option></option>を取得
'引　数：vCode			：チェック中のコード
'　　　：vAttribute		：optionの追加属性
'　　　：vForeignFlag	：海外表示フラグ
'戻り値：String
'備　考：
'履　歴：2009/08/06 LIS K.Kokubo 作成
'　　　：2009/08/21 LIS K.Kokubo vAttribute,vForeignFlag追加
'*******************************************************************************
Function htmlPrefectureOption(ByVal vCode, ByVal vAttribute, ByVal vForeignFlag)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sSQLErr

	Dim dbPrefectureCode
	Dim dbPrefectureName

	Dim sHTML
	Dim aCode
	Dim aFilter
	Dim sSelected

	sHTML = ""

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

	Do While GetRSState(oRS) = True
		dbPrefectureCode = oRS.Collect("PrefectureCode")
		dbPrefectureName = oRS.Collect("PrefectureName")

		sSelected = ""
		If UBound(Filter(aCode, dbPrefectureCode)) >= 0 Then sSelected = " selected"

		sHTML = sHTML & "<option value=""" & dbPrefectureCode & """" & sSelected & vAttribute & ">" & dbPrefectureName & "</option>"

		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	htmlPrefectureOption = sHTML
End Function
%>
