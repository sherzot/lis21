<%
'*******************************************************************************
'�T�@�v�F�s���{�������擾
'���@���FvPrefectureCode	�F�s���{���R�[�h
'�߂�l�FString
'���@�l�F
'���@���F2010/08/11 LIS K.Kokubo �쐬
'*******************************************************************************
Function getPrefectureName(ByVal vPrefectureCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sSQLErr

	sSQL = "SELECT PrefectureName FROM vw_Prefecture WHERE PrefectureCode = '" & vPrefectureCode & "';"
	flgQE = QUERYEXE(dbconn,oRS,sSQL,sSQLErr)
	If GetRSState(oRS) = True Then
		getPrefectureName = oRS.Collect("PrefectureName")
	End If
	Call RSClose(oRS)
End Function
%>
