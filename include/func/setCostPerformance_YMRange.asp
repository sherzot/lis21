<%
'*******************************************************************************
'�T�@�v�F
'���@���FrDB				�F�ڑ���DB�R�l�N�V����
'�@�@�@�FvApplicationCode	�F���O�C�������[�U�̃��C�Z���X�R�[�h
'�@�@�@�FrY1				�F[OUTPUT]�����N
'�@�@�@�FrM1				�F[OUTPUT]������
'�@�@�@�FrY2				�F[OUTPUT]����N
'�@�@�@�FrM2				�F[OUTPUT]�����
'�߂�l�FBoolean�F[True]�N���̋����ύX������Ȃ��ꍇ [False]�N���̋����ύX�����ꂽ�ꍇ
'���@�l�F
'���@���F2010/02/16 LIS K.Kokubo �쐬
'*******************************************************************************
Function setCostPerformance_YMRange(ByRef rDB,ByVal vApplicationCode,ByRef rY1,ByRef rM1,ByRef rY2,ByRef rM2)
	Dim oRS
	Dim sSQL
	Dim flgQE
	Dim sSQLErr
	'DB
	Dim dbHakouDate
	Dim dbRiyoToDate

	Dim sY1
	Dim sM1
	Dim sY2
	Dim sM2

	Dim sYM1_1
	Dim sYM2_1
	Dim sYM1_2
	Dim sYM2_2

	If IsNumber(rY1,0,False) = False Then rY1 = ""
	If IsNumber(rM1,0,False) = False Then rY2 = ""
	If IsNumber(rY2,0,False) = False Then rM1 = ""
	If IsNumber(rM2,0,False) = False Then rM2 = ""

	sYM1_1 = rY1 & Right("0"&rM1,2)
	sYM2_1 = rY2 & Right("0"&rM2,2)

	setCostPerformance_YMRange = True

	sSQL = "EXEC up_DtlNaviLicense '" & vApplicationCode & "';"
	flgQE = QUERYEXE(rDB,oRS,sSQL,sSQLErr)
	If GetRSState(oRS) = True Then
		dbHakouDate = ChkStr(oRS.Collect("HakouDate"))
		dbRiyoToDate = ChkStr(oRS.Collect("RiyoToDate"))

		sY1 = Year(dbHakouDate)
		sM1 = Month(dbHakouDate)
		sY2 = Year(dbRiyoToDate)
		sM2 = Month(dbRiyoToDate)

		sYM1_2 = sY1 & Right("0"&sM1,2)
		sYM2_2 = sY2 & Right("0"&sM2,2)

		If ChkDate8(sYM1_1&"01") = True Then
			If CLng(sYM1_1) < CLng(sYM1_2) Then
				rY1 = sY1
				rM1 = sM1
				If sYM1_1 <> "" Then setCostPerformance_YMRange = False
			End If
		ElseIf ChkDate8(sYM1_1&"01") = False Then
			rY1 = sY1
			rM1 = sM1
			If sYM1_1 <> "" Then setCostPerformance_YMRange = False
		End If

		If ChkDate8(sYM2_1&"01") = True Then
			If CLng(sYM2_1) > CLng(sYM2_2) Then
				rY2 = sY2
				rM2 = sM2
				If sYM2_1 <> "" Then setCostPerformance_YMRange = False
			End If
		ElseIf ChkDate8(sYM2_1&"01") = False Then
			rY2 = sY2
			rM2 = sM2
			If sYM2_1 <> "" Then setCostPerformance_YMRange = False
		End If
	End If
	Call RSClose(oRS)
End Function
%>
