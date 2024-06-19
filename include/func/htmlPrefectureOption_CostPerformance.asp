<%
'*******************************************************************************
'�T�@�v�F�T�|�V�X���p�s���{���ꗗ��<option></option>���擾
'���@���FvUserID	�F���O�C�������[�UID
'�@�@�@�FvCode		�F�`�F�b�N���̃R�[�h
'�@�@�@�FvAttribute	�Foption�̒ǉ�����
'�߂�l�FString
'���@�l�F
'���@���F2010/01/05 LIS K.Kokubo �쐬
'*******************************************************************************
Function htmlPrefectureOption_CostPerformance(ByVal vUserID, ByVal vCode, ByVal vAttribute)
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
	sSQL = sSQL & "/* �s���{���ꗗ */" & vbCrLf
	sSQL = sSQL & "EXEC up_LstCMPCostPerformance_Prefecture '" & vUserID & "';"
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

	htmlPrefectureOption_CostPerformance = sHTML
End Function
%>
