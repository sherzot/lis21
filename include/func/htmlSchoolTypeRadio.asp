<%
'*******************************************************************************
'�T�@�v�FB_Code �� input type="radio" �^�O��𐶐�
'���@���FvCode	�F�I�𒆂̃R�[�h
'�@�@�@�FvCols	�F��
'�@�@�@�FvName	�Fname�����l
'�@�@�@�FvJHFlag�F���w�Z�̕\���E��\���t���O ["0"]��\�� ["1"]�\��
'�߂�l�FString
'���@���F2011/02/11 LIS K.Kokubo �쐬
'*******************************************************************************
Function htmlSchoolTypeRadio(ByVal vCode,ByVal vCols,ByVal vName,ByVal vJHFlag)
	Dim sSQL,oRS,flgQE,sSQLErr
	Dim dbSchoolTypeCode,dbSchoolTypeName
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
	sSQL = sSQL & "SELECT SchoolTypeCode, SchoolTypeName FROM vw_SchoolType"
	If vJHFlag = "0" Then sSQL = sSQL & " WHERE SchoolTypeCode > '001'"
	sSQL = sSQL & ";"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sSQLErr)
	Do While GetRSState(oRS) = True
		dbSchoolTypeCode = oRS.Collect("SchoolTypeCode")
		dbSchoolTypeName = oRS.Collect("SchoolTypeName")

		sChecked = ""
		If UBound(Filter(aValue, dbSchoolTypeCode)) >= 0 Then sChecked = " checked"

		If idx Mod CInt(vCols) = 1 Then sHTML = sHTML & "<tr>"

		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<label>"
		sHTML = sHTML & "<input name=""" & vName & """ type=""radio"" value=""" & dbSchoolTypeCode & """" & sChecked & ">" & dbSchoolTypeName
		sHTML = sHTML & "</label>"
		sHTML = sHTML & "</td>"

		oRS.MoveNext

		If GetRSState(oRS) = False Or idx Mod CInt(vCols) = 0 Then sHTML = sHTML & "</tr>"

		idx = idx + 1
	Loop
	Call RSClose(oRS)

	sHTML = sHTML & "</tbody>"
	sHTML = sHTML & "</table>"

	htmlSchoolTypeRadio = sHTML
End Function
%>
