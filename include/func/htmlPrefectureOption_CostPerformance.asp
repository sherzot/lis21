<%
'*******************************************************************************
'概　要：サポシス利用都道府県一覧の<option></option>を取得
'引　数：vUserID	：ログイン中ユーザID
'　　　：vCode		：チェック中のコード
'　　　：vAttribute	：optionの追加属性
'戻り値：String
'備　考：
'履　歴：2010/01/05 LIS K.Kokubo 作成
'*******************************************************************************
Function htmlPrefectureOption_CostPerformance(ByVal vUserID, ByVal vCode, ByVal vAttribute)
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
	sSQL = sSQL & "/* 都道府県一覧 */" & vbCrLf
	sSQL = sSQL & "EXEC up_LstCMPCostPerformance_Prefecture '" & vUserID & "';"
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

	htmlPrefectureOption_CostPerformance = sHTML
End Function
%>
