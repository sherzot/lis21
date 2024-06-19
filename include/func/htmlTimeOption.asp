<%
'*******************************************************************************
'�T�@�v�F�s���{���ꗗ��<option></option>���擾
'���@���FvTime		�F�I�𒆂̎���
'�@�@�@�FvSepMinute	�F���̋�؂�...15�̏ꍇ00,15,30,45�A30�̏ꍇ00,30�A60���������Ƃ��Ɋ���؂�鐔�����
'�߂�l�FString
'���@�l�F
'���@���F2009/08/31 LIS K.Kokubo �쐬
'*******************************************************************************
Function htmlTimeOption(ByVal vTime, ByVal vSepMinute)
	Dim sHTML
	Dim aMinute()
	Dim iSep
	Dim idx
	Dim idx2
	Dim sTime
	Dim sSelected

	If 60 Mod CInt(vSepMinute) <> 0 Then Exit Function

	iSep = 60 / CInt(vSepMinute)
	ReDim aMinute(iSep - 1)
	For idx = 0 To iSep - 1
		aMinute(idx) = Right("0" & CInt(vSepMinute) * idx, 2)
	Next

	sHTML = ""
	For idx = 0 To 23
		For idx2 = 0 To UBound(aMinute)
			sSelected = ""

			sTime = Right("0" & idx, 2) & aMinute(idx2)
			If sTime = vTime Then sSelected = " selected"

			sHTML = sHTML & "<option value=""" & sTime & """" & sSelected & ">" & Right("0" & idx, 2) & ":" & aMinute(idx2) & "</option>"
		Next
	Next

	htmlTimeOption = sHTML
End Function
%>
