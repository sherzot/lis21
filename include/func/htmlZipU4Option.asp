<%
'*******************************************************************************
'�T�@�v�F�s���{���ꗗ�̃`�F�b�N�{�b�N�X���擾
'���@���FvCode	�F�`�F�b�N���̃R�[�h
'�@�@�@�FvCols	�F��s������̗�
'�@�@�@�FvName	�Finput��name�����l
'�߂�l�FString
'���@�l�F
'���@���F2009/08/05 LIS K.Kokubo �쐬
'*******************************************************************************
Function htmlZipU4Option(ByVal vPrefectureCode, ByVal vCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sSQLErr

	Dim dbZipCode
	Dim dbAddr

	Dim sHTML
	Dim aCode
	Dim aFilter
	Dim sSelected

	sHTML = ""

	sSQL = ""
	sSQL = sSQL & "/* �s���{���ꗗ */" & vbCrLf
	sSQL = sSQL & "EXEC up_LstB_Zip_U4 '" & vPrefectureCode & "';"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sSQLErr)

	aCode = Split(vCode, ",")

	Do While GetRSState(oRS) = True
		dbZipCode = oRS.Collect("ZipCode")
		dbAddr = oRS.Collect("Addr")

		sSelected = ""
		If UBound(Filter(aCode, dbZipCode)) >= 0 Then sSelected = " selected"

		sHTML = sHTML & "<option value=""" & dbZipCode & """" & sSelected & ">" & Left(dbZipCode, 3) & "-" & Right(dbZipCode, 1) & "XXX&nbsp;(" & dbAddr & ")</option>"

		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	htmlZipU4Option = sHTML
End Function
%>
