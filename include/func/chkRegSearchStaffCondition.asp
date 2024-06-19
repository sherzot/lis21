<%
'*******************************************************************************
'概　要：求職者検索条件の保存の可否
'引　数：vCnt	：現在の保存件数
'戻り値：Boolean
'備　考：
'更　新：2009/06/25 LIS K.Kokubo 作成
'*******************************************************************************
Function chkRegSearchStaffCondition(ByVal vCnt)
	chkRegSearchStaffCondition = True
	If vCnt >= 3 Then chkRegSearchStaffCondition = False
End Function
%>
