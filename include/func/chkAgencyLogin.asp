<%
'*******************************************************************************
'概　要：代理店ログインチェック
'引　数：
'戻り値：Boolean
'備　考：
'更　新：2010/03/29 LIS K.Kokubo 作成
'*******************************************************************************
Function chkAgencyLogin(ByVal vAgencyCode,ByVal vBranchSeq,ByVal vPwd)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sSQLErr

	chkAgencyLogin = False

	sSQL = "EXEC up_ChkAGC_Login '" & vAgencyCode & "','" & vBranchSeq & "','" & vPwd & "';"
	flgQE = QUERYEXE(dbconn,oRS,sSQL,sSQLErr)
	If GetRSState(oRS) = True Then
		chkAgencyLogin = True

		Session("agencycode") = vAgencyCode
		Session("agencybranch") = vBranchSeq
	End If
	Call RSClose(oRS)
End Function
%>
