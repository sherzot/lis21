<%
'*******************************************************************************
'�T�@�v�F�Z�p���[�g������ǉ����ĕ����������
'���@���FvStr1	�F��������镶����
'�@�@�@�FvStr2	�F�������镶����
'�@�@�@�FvSep	�F��������镶���񂪋󕶎��Ŗ����ꍇ�ɁA���������؂蕶��(�Z�p���[�g)
'�߂�l�FString
'���@�l�F
'���@���F2011/02/28 LIS K.Kokubo �쐬
'*******************************************************************************
Function getStrJoinSep(ByVal vStr1,ByVal vStr2,ByVal vSep)
	getStrJoinSep = vStr1

	If getStrJoinSep <> "" Then getStrJoinSep = getStrJoinSep & vSep
	getStrJoinSep = getStrJoinSep & vStr2
End Function
%>
