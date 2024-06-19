<%
'*******************************************************************************
'概　要：採用改善サポートシステムの媒体名一覧の<option></option>を取得
'引　数：vMedName		：選択中の媒体名
'　　　：vAttribute		：optionの追加属性
'戻り値：String
'備　考：
'履　歴：2009/10/29 LIS K.Kokubo 作成
'*******************************************************************************
Function htmlMedNameOption(ByVal vMedName, ByVal vAttribute)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sSQLErr

	Dim dbMedName

	Dim sHTML
	Dim aMedName
	Dim aFilter
	Dim sSelected

	sHTML = ""

	If vAttribute <> "" Then vAttribute = " " & vAttribute

	sSQL = ""
	sSQL = sSQL & "/* 改善サポートシステム媒体一覧 */" & vbCrLf
	sSQL = sSQL & "EXEC up_LstMedName '';"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sSQLErr)

	aMedName = Split(vMedName, ",")
	Do While GetRSState(oRS) = True
		dbMedName = oRS.Collect("MedName")

		sSelected = ""
		If UBound(Filter(aMedName, dbMedName)) >= 0 Then sSelected = " selected"

		sHTML = sHTML & "<option value=""" & dbMedName & """" & vAttribute & sSelected & ">" & dbMedName & "</option>"

		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	htmlMedNameOption = sHTML
End Function
%>
