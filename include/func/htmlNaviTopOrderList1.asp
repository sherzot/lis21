<%
'*******************************************************************************
'�T�@�v�F�i�r�̂s�n�o�y�[�W�ɕ\�����郍�O�C���ς݋��E�҂̋��l�������ʈꗗ�i�ۑ����Ă����������j
'���@���F
'�o�@�́F
'�߂�l�FString
'���@�l�F�ۑ����Ă��錟�������̔ԍ� seq=1
'���@���F2010/11/06 LIS K.Kokubo �쐬
'*******************************************************************************
Function htmlNaviTopOrderList1(ByRef rDB,ByVal vUserID)
	Dim oRS,oRS2,sSQL,flgQE,sSQLErr
	Dim dbSearchName,dbSearchParam
	Dim oSOC
	Dim iCnt

	sSQL = "EXEC up_LstP_SearchOrderCondition '" & vUserID & "';"
	flgQE = QUERYEXE(rDB,oRS,sSQL,sSQLErr)
	If GetRSState(oRS) = True Then
		Set oRS.ActiveConnection = Nothing

		dbSearchName = oRS.Collect("SearchName")
		dbSearchParam = oRS.Collect("SearchParam")

		Set oSOC = New clsSearchOrderCondition
		oSOC.SetData_Param(dbSearchParam)
		sSQL = oSOC.GetSQLOrderSearchDetail()
		flgQE = QUERYEXE(rDB,oRS2,sSQL,sSQLErr)
		If GetRSState(oRS2) = True Then
			Set oRS2.ActiveConnection = Nothing

			iCnt = oRS2.RecordCount

			sHTML = htmlOrderListLine(rDB,oRS2,5)
			If iCnt > 5 Then
				sHTML = sHTML & "<p style=""text-align:right;""><a href=""" & HTTP_CURRENTURL & "order/order_list.asp?" & Replace(dbSearchParam,"&","&amp;") & """>" & _
					"&gt;&gt;�����ƌ������ʂ�����" & _
					"</a></p>"
			End If
		Else
			sHTML = "<p>�ۑ��������l�̌��������Ƀ}�b�`���邨���l��������܂���ł����B</p>"
		End If
	Else
		sHTML = "<p>���l�̌����������ۑ�����Ă��܂���B" & _
			"���l�̌���������ۑ�����ɂ́A���l�̌������ʈꗗ�Łu���̌���������ۑ�����v���N���b�N���܂��B" & _
			"���l�̌�����<a href=""" & HTTP_CURRENTURL & "order/order_search_detail.asp"">�R�`��</a>����ǂ����B</p>"
	End If
	Call RSClose(oRS)

	htmlNaviTopOrderList1 = sHTML
End Function
%>
