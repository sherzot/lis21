<%
'*******************************************************************************
'�T�@�v�FB_Code �� input type="radio" �^�O��𐶐�
'���@���FvCode	�F�I�𒆂̃R�[�h
'�@�@�@�FvName	�Fname�����l
'�@�@�@�FvJHFlag�F���w�Z�̕\���E��\���t���O ["0"]��\�� ["1"]�\��
'�߂�l�FString
'���@���F2009/08/06 LIS K.Kokubo �쐬
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
