<%
'*******************************************************************************
'�T�@�v�F�s���{���ꗗ�̃`�F�b�N�{�b�N�X���擾
'���@���FvCompanyCode	�F��ƃR�[�h
'�@�@�@�FvCode			�F�`�F�b�N���̃R�[�h
'�@�@�@�FvCols			�F��s������̗�
'�@�@�@�FvName			�Finput��name�����l
'�߂�l�FString
'���@�l�F
'���@���F2009/08/05 LIS K.Kokubo �쐬
'*******************************************************************************
Function htmlWorkingTypeCompanyUseRadio(ByVal vCompanyCode, ByVal vCode, ByVal vCols, ByVal vName)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sSQLErr

	Dim dbWorkingTypeCode
	Dim dbWorkingTypeName

	Dim sHTML
	Dim aCode
	Dim aFilter
	Dim sChecked
	Dim fWidth
	Dim idx

	sHTML = ""
	If vCols > 0 Then fWidth = Round(100 / CInt(vCols), 1)

	sSQL = ""
	sSQL = sSQL & "/* ��Ƃ����p�ł���Ζ��`�Ԉꗗ */" & vbCrLf
	sSQL = sSQL & "EXEC up_LstWorkingType_CompanyUse '" & vCompanyCode & "';"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sSQLErr)

	aCode = Split(vCode, ",")

	sHTML = sHTML & "<table border=""0"" style=""width:100%;"">" & vbCrLf
	If CInt(vCols) > 0 Then
		sHTML = sHTML & "<colgroup>"
		For idx = 0 To CInt(vCols) - 1
			sHTML = sHTML & "<col style=""width:" & fWidth & "%;""></col>"
		Next
		sHTML = sHTML & "</colgroup>"
	End If
	sHTML = sHTML & "<tbody>"

	idx = 0
	Do While GetRSState(oRS) = True
		dbWorkingTypeCode = oRS.Collect("WorkingTypeCode")
		dbWorkingTypeName = oRS.Collect("WorkingTypeName")
		If dbWorkingTypeCode = "005" Then
			dbWorkingTypeName = "�p�[�g�E�A���o�C�g"
		ElseIf dbWorkingTypeCode = "006" Then
			dbWorkingTypeName = "�r�n�g�n"
		End If

		If idx = 0 Then
			sHTML = sHTML & "<tr>"
		End If

		sChecked = ""
		If UBound(Filter(aCode, dbWorkingTypeCode)) >= 0 Then sChecked = " checked"

		sHTML = sHTML & "<td style=""padding:0px;border-width:0px;"">"
		sHTML = sHTML & "<label><input name=""" & vName & """ type=""radio"" value=""" & dbWorkingTypeCode & """" & sChecked & ">" & dbWorkingTypeName & "</label>"
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

	htmlWorkingTypeCompanyUseRadio = sHTML
End Function
%>
