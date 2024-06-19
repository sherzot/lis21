<%
'*******************************************************************************
'概　要：B_Code の input type="radio" タグ句を生成
'引　数：vCode	：選択中のコード
'　　　：vName	：name属性値
'　　　：vJHFlag：中学校の表示・非表示フラグ ["0"]非表示 ["1"]表示
'戻り値：String
'履　歴：2009/08/06 LIS K.Kokubo 作成
'*******************************************************************************
Function htmlSchoolTypeCheckbox(ByVal vCode, ByVal vName, ByVal vJHFlag)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sSQLErr

	Dim dbSchoolTypeCode
	Dim dbSchoolTypeName

	Dim sHTML
	Dim aValue
	Dim sChecked

	sHTML = ""

	aValue = Split(Replace(vCode, " ", ""), ",")

	sSQL = ""
	sSQL = sSQL & "SELECT SchoolTypeCode, SchoolTypeName FROM vw_SchoolType"
	If vJHFlag = "0" Then sSQL = sSQL & " WHERE SchoolTypeCode > '001'"
	sSQL = sSQL & ";"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sSQLErr)
	Do While GetRSState(oRS) = True
		dbSchoolTypeCode = oRS.Collect("SchoolTypeCode")
		dbSchoolTypeName = oRS.Collect("SchoolTypeName")

		sChecked = ""
		If UBound(Filter(aValue, dbSchoolTypeCode)) >= 0 Then sChecked = " checked"

		sHTML = sHTML & "<label>"
		sHTML = sHTML & "<input name=""" & vName & """ type=""checkbox"" value=""" & dbSchoolTypeCode & """" & sChecked & ">" & dbSchoolTypeName
		sHTML = sHTML & "</label>"

		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	htmlSchoolTypeCheckbox = sHTML
End Function
%>
