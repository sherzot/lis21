<%
'*******************************************************************************
'�T�@�v�F���l�[�̋Ζ��`�Ԉꗗ�̃��a�I�{�^�����擾
'���@���FvOrderCode	�F���R�[�h
'�@�@�@�FvCode		�F�`�F�b�N���̃R�[�h
'�@�@�@�FvCols		�F��s������̗�
'�@�@�@�FvName		�Finput��name�����l
'�߂�l�FString
'���@�l�F
'���@���F2009/08/05 LIS K.Kokubo �쐬
'*******************************************************************************
Function htmlWorkingTypeOrderUseRadio(ByVal vOrderCode, ByVal vCode, ByVal vCols, ByVal vName)
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
	sSQL = sSQL & "/* ���l�[�̋Ζ��`�Ԉꗗ */" & vbCrLf
	sSQL = sSQL & "EXEC sp_GetDataWorkingType '" & vOrderCode & "';"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sSQLErr)

	aCode = Split(vCode, ",")

	If CInt(vCols) > 0 Then
		sHTML = sHTML & "<table border=""0"" style=""width:100%;"">" & vbCrLf
		sHTML = sHTML & "<colgroup>"
		For idx = 0 To CInt(vCols) - 1
			sHTML = sHTML & "<col style=""width:" & fWidth & "%;""></col>"
		Next
		sHTML = sHTML & "</colgroup>"
		sHTML = sHTML & "<tbody>"
	End If

	idx = 0
	Do While GetRSState(oRS) = True
		dbWorkingTypeCode = oRS.Collect("WorkingTypeCode")
		dbWorkingTypeName = oRS.Collect("WorkingTypeName")
		If dbWorkingTypeCode = "005" Then
			dbWorkingTypeName = "�p�[�g�E�A���o�C�g"
		ElseIf dbWorkingTypeCode = "006" Then
			dbWorkingTypeName = "�r�n�g�n"
		End If

		If vCols > 0 Then
			If idx = 0 Then
				sHTML = sHTML & "<tr>"
			End If
		End If

		sChecked = ""
		If UBound(Filter(aCode, dbWorkingTypeCode)) >= 0 Then sChecked = " checked"

		If vCols > 0 Then sHTML = sHTML & "<td style=""padding:0px;border-width:0px;"">"
		sHTML = sHTML & "<label><input name=""" & vName & """ type=""radio"" value=""" & dbWorkingTypeCode & """" & sChecked & ">" & dbWorkingTypeName & "</label>"
		If vCols > 0 Then sHTML = sHTML & "</td>"
		If vCols = 0 Then
			If sHTML <> "" Then sHTML = sHTML & "&nbsp;"
		End If

		oRS.MoveNext

		If vCols > 0 Then
			If GetRSState(oRS) = False Or idx = CInt(vCols) - 1 Then
				sHTML = sHTML & "</tr>"
				idx = 0
			Else
				idx = idx + 1
			End If
		End If
	Loop
	Call RSClose(oRS)

	If vCols > 0 Then
		sHTML = sHTML & "</tbody>"
		sHTML = sHTML & "</table>"
	End If

	htmlWorkingTypeOrderUseRadio = sHTML
End Function
%>
