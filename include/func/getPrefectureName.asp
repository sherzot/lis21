<%
'*******************************************************************************
'概　要：都道府県名を取得
'引　数：vPrefectureCode	：都道府県コード
'戻り値：String
'備　考：
'履　歴：2010/08/11 LIS K.Kokubo 作成
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
