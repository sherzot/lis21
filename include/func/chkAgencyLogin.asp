<%
'*******************************************************************************
'�T�@�v�F�㗝�X���O�C���`�F�b�N
'���@���F
'�߂�l�FBoolean
'���@�l�F
'�X�@�V�F2010/03/29 LIS K.Kokubo �쐬
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
