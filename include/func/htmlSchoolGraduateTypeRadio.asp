<%
'*******************************************************************************
'�T�@�v�FB_Code �� input type="radio" �^�O��𐶐�
'���@���FvCode	�F�I�𒆂̃R�[�h
'�@�@�@�FvCols	�F��
'�@�@�@�FvName	�Fname�����l
'�߂�l�FString
'���@���F2011/01/21 LIS K.Kokubo �쐬
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
