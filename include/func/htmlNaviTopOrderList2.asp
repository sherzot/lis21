<%
'*******************************************************************************
'�T�@�v�F�i�r�̂s�n�o�y�[�W�ɕ\�����郍�O�C���ς݋��E�҂̋��l�������ʈꗗ�i���������j
'���@���F
'�o�@�́F
'�߂�l�FString
'���@�l�F
'���@���F2010/10/29 LIS K.Kokubo �쐬
'*******************************************************************************
Function htmlNaviTopOrderList2(ByRef rDB,ByVal vUserID)
	Dim oRS,oRS2,sSQL,flgQE,sSQLErr
	Dim dbOrderCode,dbJobTypeDetail
	Dim dbWorkingPlacePrefectureName,dbWorkingPlaceCity,dbWorkingTypeName
	Dim dbYearlyIncomeMin,dbYearlyIncomeMax,dbMonthlyIncomeMin,dbMonthlyIncomeMax,dbDailyIncomeMin,dbDailyIncomeMax,dbHourlyIncomeMin,dbHourlyIncomeMax
	Dim sHTML,sWP,sWT,sSalary
	Dim idx,idx2

	sSQL = "EXEC up_SearchOrderAuto '" & vUserID & "','';"
	flgQE = QUERYEXE(rDB,oRS,sSQL,sSQLErr)
	If GetRSState(oRS) = True Then
		Set oRS.ActiveConnection = Nothing

		sHTML = htmlOrderListLine(rDB,oRS,5)
	Else
		sHTML = "<p>���Ȃ��̊�]�����Ƀ}�b�`���鋁�l��������܂���ł����B</p>"
	End If

	htmlNaviTopOrderList2 = sHTML
End Function
%>
