<%
'*******************************************************************************
'概　要：B_Code の input type="radio" タグ句を生成
'引　数：vCode	：選択中のコード
'　　　：vCols	：列数
'　　　：vName	：name属性値
'戻り値：String
'履　歴：2011/01/21 LIS K.Kokubo 作成
'*******************************************************************************
Function htmlSchoolGraduateTypeRadio(ByVal vCode,ByVal vCols,ByVal vName)
	Dim sSQL,oRS,flgQE,sSQLErr
	Dim dbGraduateTypeCode,dbGraduateTypeName

	Dim sHTML,aValue,sChecked
	Dim idx,fWidth

	sHTML = ""

	aValue = Split(Replace(ChkStr(vCode), " ", ""), ",")


	fWidth = Round(100 / CInt(vCols), 1)
	sHTML = sHTML & "<table border=""0"" style=""width:100%;"">" & vbCrLf
	sHTML = sHTML & "<colgroup>"
	For idx = 0 To CInt(vCols) - 1
		sHTML = sHTML & "<col style=""width:" & fWidth & "%;""></col>"
	Next
	sHTML = sHTML & "</colgroup>"
	sHTML = sHTML & "<tbody>"


	idx = 1

	sSQL = ""
	sSQL = sSQL & "SELECT GraduateTypeCode, GraduateTypeName FROM vw_SchoolGraduateType;"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sSQLErr)
	Do While GetRSState(oRS) = True
		dbGraduateTypeCode = oRS.Collect("GraduateTypeCode")
		dbGraduateTypeName = oRS.Collect("GraduateTypeName")

		sChecked = ""
		If UBound(Filter(aValue, dbGraduateTypeCode)) >= 0 Then sChecked = " checked"

		If idx Mod CInt(vCols) = 1 Then sHTML = sHTML & "<tr>"

		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<label>"
		sHTML = sHTML & "<input name=""" & vName & """ type=""radio"" value=""" & dbGraduateTypeCode & """" & sChecked & ">" & dbGraduateTypeName
		sHTML = sHTML & "</label>&nbsp;"
		sHTML = sHTML & "</td>"

		oRS.MoveNext

		If GetRSState(oRS) = False Or idx Mod CInt(vCols) = 0 Then sHTML = sHTML & "</tr>"

		idx = idx + 1
	Loop
	Call RSClose(oRS)


	sHTML = sHTML & "</tbody>"
	sHTML = sHTML & "</table>"


	htmlSchoolGraduateTypeRadio = sHTML
End Function
%>
