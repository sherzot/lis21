<%
'*******************************************************************************
'�T�@�v�F�㗝�X���_�ꗗ��<option></option>���擾
'���@���FvAgencyCode�F�㗝�X�R�[�h
'�@�@�@�FvBranchSeq	�F�㗝�X���_�ԍ�
'�@�@�@�FvCode		�F�`�F�b�N���S���Ҕԍ�
'�@�@�@�FvAttribute	�Foption�̒ǉ�����
'�߂�l�FString
'���@�l�F
'���@���F2010/03/17 LIS K.Kokubo �쐬
'*******************************************************************************
Function htmlAGCChargeOption(ByVal vAgencyCode, ByVal vBranchSeq, ByVal vCode, ByVal vAttribute)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sSQLErr

	Dim dbChargeSeq
	Dim dbPersonName

	Dim sHTML
	Dim aCode
	Dim aFilter
	Dim sSelected

	sHTML = ""

	If vAttribute <> "" Then vAttribute = " " & vAttribute

	sSQL = ""
	sSQL = sSQL & "/* �㗝�X���_�ꗗ */" & vbCrLf
	sSQL = sSQL & "SELECT ChargeSeq,PersonName FROM AGCCharge WHERE AgencyCode = '" & ChkStr(vAgencyCode) & "' AND BranchSeq = '" & ChkStr(vBranchSeq) & "';"
	flgQE = QUERYEXE(dbconn,oRS,sSQL,sSQLErr)
	aCode = Split(ChkStr(vCode), ",")
	Do While GetRSState(oRS) = True
		dbChargeSeq = oRS.Collect("ChargeSeq")
		dbPersonName = oRS.Collect("PersonName")

		sSelected = ""
		If UBound(Filter(aCode, dbChargeSeq)) >= 0 Then sSelected = " selected"

		sHTML = sHTML & "<option value=""" & dbChargeSeq & """" & vAttribute & sSelected & ">" & dbPersonName & "</option>"

		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	htmlAGCChargeOption = sHTML
End Function
%>
