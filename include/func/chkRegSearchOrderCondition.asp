<%
'*******************************************************************************
'概　要：求職者検索条件の保存の可否
'引　数：vCnt	：現在の保存件数
'戻り値：Boolean
'備　考：
'更　新：2009/06/25 LIS K.Kokubo 作成
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
