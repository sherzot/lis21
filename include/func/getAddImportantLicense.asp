<%
'*******************************************************************************
'�T�@�v�F���E�҂��ǉ��o�^�����d�v���i�����擾
'���@���F
'�߂�l�FString
'���@�l�F
'���@���F2010/08/25 LIS K.Kokubo �쐬
'*******************************************************************************
Function getAddImportantLicense(ByRef rBefore,ByRef rAfter)
	Dim idx
	Dim tmpAry,sImportant

	getAddImportantLicense = ""
	tmpAry = rBefore.Items
	sImportant = ""

	idx = 1
	Do While rBefore.Exists("LicenseName"&idx) = True Or rAfter.Exists("LicenseName"&idx) = True
		If UBound(tmpAry) > 0 And rAfter("LicenseName" & idx) <> "" Then
			If rAfter("LicenseName" & idx) = "���p���Z�p��" And UBound(Filter(tmpAry,"���p���Z�p��")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "���p���Z�p��"
			ElseIf rAfter("LicenseName" & idx) = "�h�s�X�g���e�W�X�g" And UBound(Filter(tmpAry,"�h�s�X�g���e�W�X�g")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "�h�s�X�g���e�W�X�g"
			ElseIf rAfter("LicenseName" & idx) = "�v���W�F�N�g�}�l�[�W��" And UBound(Filter(tmpAry,"�v���W�F�N�g�}�l�[�W��")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "�v���W�F�N�g�}�l�[�W��"
			ElseIf rAfter("LicenseName" & idx) = "�V�X�e���A�[�L�e�N�g" And UBound(Filter(tmpAry,"�V�X�e���A�[�L�e�N�g")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "�V�X�e���A�[�L�e�N�g"
			ElseIf rAfter("LicenseName" & idx) = "�h�s�T�[�r�X�}�l�[�W��" And UBound(Filter(tmpAry,"�h�s�T�[�r�X�}�l�[�W��")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "�h�s�T�[�r�X�}�l�[�W��"
			ElseIf rAfter("LicenseName" & idx) = "���Z�L�����e�B�X�y�V�����X�g" And UBound(Filter(tmpAry,"���Z�L�����e�B�X�y�V�����X�g")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "���Z�L�����e�B�X�y�V�����X�g"
			ElseIf rAfter("LicenseName" & idx) = "CCNP" And UBound(Filter(tmpAry,"CCNP")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "CCNP"
			ElseIf rAfter("LicenseName" & idx) = "LPIC Level 3" And UBound(Filter(tmpAry,"LPIC Level 3")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "LPIC Level 3"
			ElseIf rAfter("LicenseName" & idx) = "�Ō�t" And UBound(Filter(tmpAry,"�Ō�t")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "�Ō�t"
			ElseIf rAfter("LicenseName" & idx) = "�P�A�}�l�[�W���[�i���x�������j" And UBound(Filter(tmpAry,"�P�A�}�l�[�W���[�i���x�������j")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "�P�A�}�l�[�W���[�i���x�������j"
			ElseIf rAfter("LicenseName" & idx) = "��܎t" And UBound(Filter(tmpAry,"��܎t")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "��܎t"
			ElseIf rAfter("LicenseName" & idx) = "��t" And UBound(Filter(tmpAry,"��t")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "��t"
			ElseIf rAfter("LicenseName" & idx) = "�Տ������Z�t" And UBound(Filter(tmpAry,"�Տ������Z�t")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "�Տ������Z�t"
			ElseIf rAfter("LicenseName" & idx) = "MR�F�莑�i" And UBound(Filter(tmpAry,"MR�F�莑�i")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "MR�F�莑�i"
			ElseIf rAfter("LicenseName" & idx) = "�ۈ�m" And UBound(Filter(tmpAry,"�ۈ�m")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "�ۈ�m"
			ElseIf rAfter("LicenseName" & idx) = "�h�{�m" And UBound(Filter(tmpAry,"�h�{�m")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "�h�{�m"
			ElseIf rAfter("LicenseName" & idx) = "�y�Ō�t" And UBound(Filter(tmpAry,"�y�Ō�t")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "�y�Ō�t"
			ElseIf rAfter("LicenseName" & idx) = "��앟���m" And UBound(Filter(tmpAry,"��앟���m")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "��앟���m"
			ElseIf rAfter("LicenseName" & idx) = "���w�Ö@�m" And UBound(Filter(tmpAry,"���w�Ö@�m")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "���w�Ö@�m"
			ElseIf rAfter("LicenseName" & idx) = "��ƗÖ@�m" And UBound(Filter(tmpAry,"��ƗÖ@�m")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "��ƗÖ@�m"
			ElseIf rAfter("LicenseName" & idx) = "�ی��t" And UBound(Filter(tmpAry,"�ی��t")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "�ی��t"
			End If
		End If

		idx = idx + 1
	Loop

	getAddImportantLicense = sImportant
End Function

Function setDicLicense(ByRef rDB,ByVal vUserCode,ByRef rDic)
	Dim sSQL,oRS,flgQE,sSQLErr
	Dim idx

	setDicLicense = False
	Set rDic = Server.CreateObject("scripting.dictionary")

	sSQL = "sp_GetDataLicense '" & vUserCode & "'"
	flgQE = QUERYEXE(rDB,oRS,sSQL,sSQLErr)
	If GetRSState(oRS) = True Then
		Set oRS.ActiveConnection = Nothing
		idx = 1
		Do While GetRSState(oRS)
			Call rDic.Add("Code" & idx, oRS.Collect("GroupCode") & oRS.Collect("CategoryCode") & oRS.Collect("Code"))
			Call rDic.Add("LicenseName" & idx, oRS.Collect("LicenseName"))
			Call rDic.Add("LicenseNameDsp" & idx, oRS.Collect("LicenseNameDsp"))
			Call rDic.Add("GetDay" & idx, ChkStr(oRS.Collect("GetDay")))

			idx = idx + 1
			oRS.MoveNext
		Loop
		setDicLicense = True
	End If
	Call RSClose(oRS)
End Function
%>
