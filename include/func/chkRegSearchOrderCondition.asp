<%
'*******************************************************************************
'�T�@�v�F���E�Ҍ��������̕ۑ��̉�
'���@���FvCnt	�F���݂̕ۑ�����
'�߂�l�FBoolean
'���@�l�F
'�X�@�V�F2009/06/25 LIS K.Kokubo �쐬
'*******************************************************************************
Function chkRegSearchOrderCondition(ByRef rDB,ByVal vStaffCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sSQLErr

	chkRegSearchOrderCondition = False

	sSQL = "EXEC up_ChkP_SearchOrderCondition '" & vStaffCode & "';"
	flgQE = QUERYEXE(rDB,oRS,sSQL,sSQLErr)
	If GetRSState(oRS) = True Then
		If oRS.Collect("Flag") = "1" Then chkRegSearchOrderCondition = True
	End If
End Function
%>
