<%
'*******************************************************************************
'�T�@�v�F���E�Ҍ��������̕ۑ��̉�
'���@���FvCnt	�F���݂̕ۑ�����
'�߂�l�FBoolean
'���@�l�F
'�X�@�V�F2009/06/25 LIS K.Kokubo �쐬
'*******************************************************************************
Function chkRegSearchStaffCondition(ByVal vCnt)
	chkRegSearchStaffCondition = True
	If vCnt >= 3 Then chkRegSearchStaffCondition = False
End Function
%>
