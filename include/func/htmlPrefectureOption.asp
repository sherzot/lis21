<%
'*******************************************************************************
'�T�@�v�F�s���{���ꗗ��<option></option>���擾
'���@���FvCode			�F�`�F�b�N���̃R�[�h
'�@�@�@�FvAttribute		�Foption�̒ǉ�����
'�@�@�@�FvForeignFlag	�F�C�O�\���t���O
'�߂�l�FString
'���@�l�F
'���@���F2009/08/06 LIS K.Kokubo �쐬
'�@�@�@�F2009/08/21 LIS K.Kokubo vAttribute,vForeignFlag�ǉ�
'*******************************************************************************
Function htmlPrefectureOption(ByVal vCode, ByVal vAttribute, ByVal vForeignFlag)
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
	If vForeignFlag = "1" Then
		sSQL = sSQL & "/* �s���{��(�C�O�܂�)�ꗗ */" & vbCrLf
		sSQL = sSQL & "EXEC up_LstPrefectureAll;"
	Else
		sSQL = sSQL & "/* �s���{���ꗗ */" & vbCrLf
		sSQL = sSQL & "EXEC up_LstPrefecture;"
	End If
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

	htmlPrefectureOption = sHTML
End Function
%>
