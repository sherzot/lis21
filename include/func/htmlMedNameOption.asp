<%
'*******************************************************************************
'�T�@�v�F�̗p���P�T�|�[�g�V�X�e���̔}�̖��ꗗ��<option></option>���擾
'���@���FvMedName		�F�I�𒆂̔}�̖�
'�@�@�@�FvAttribute		�Foption�̒ǉ�����
'�߂�l�FString
'���@�l�F
'���@���F2009/10/29 LIS K.Kokubo �쐬
'*******************************************************************************
Function htmlMedNameOption(ByVal vMedName, ByVal vAttribute)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sSQLErr

	Dim dbMedName

	Dim sHTML
	Dim aMedName
	Dim aFilter
	Dim sSelected

	sHTML = ""

	If vAttribute <> "" Then vAttribute = " " & vAttribute

	sSQL = ""
	sSQL = sSQL & "/* ���P�T�|�[�g�V�X�e���}�̈ꗗ */" & vbCrLf
	sSQL = sSQL & "EXEC up_LstMedName '';"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sSQLErr)

	aMedName = Split(vMedName, ",")
	Do While GetRSState(oRS) = True
		dbMedName = oRS.Collect("MedName")

		sSelected = ""
		If UBound(Filter(aMedName, dbMedName)) >= 0 Then sSelected = " selected"

		sHTML = sHTML & "<option value=""" & dbMedName & """" & vAttribute & sSelected & ">" & dbMedName & "</option>"

		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	htmlMedNameOption = sHTML
End Function
%>
