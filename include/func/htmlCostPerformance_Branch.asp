<%
'*******************************************************************************
'�T�@�v�F�̗p���P�T�|�[�g�V�X�e���̕���E�X�ܕʔ�p�Ό���TABLE���擾
'���@���FrDB	�F�ڑ���DB�R�l�N�V����
'�@�@�@�FvUserID�F���O�C�������[�UID
'�@�@�@�FvYM1	�F�W�v���ԉ����N��
'�@�@�@�FvYM2	�F�W�v���ԏ���N��
'�߂�l�FString
'���@�l�F
'���@���F2010/03/10 LIS K.Kokubo �쐬
'*******************************************************************************
Function htmlCostPerformance_Branch(ByRef rDB, ByRef rRS, ByVal vUserID, ByVal rCP, ByVal vPageSize, ByVal vPage, ByVal vSort, ByVal vBranchName)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sSQLErr
	'DB
	Dim dbCompanyCode
	Dim dbBranchSeq
	Dim dbBranchName
	Dim dbMedName
	Dim dbCost
	Dim dbAdoptNumPlan
	Dim dbAdoptNumResult
	Dim dbUnitCost
	Dim dbAdoptNumPlanPeriod

	Dim iCost
	Dim iUnitCost
	Dim iAdoptNumPlan
	Dim iAdoptNum

	Dim idx
	Dim sHref
	Dim sFlag
	Dim sHTML
	Dim sFilter

	Dim aBN

	sHTML = ""

	If GetRSState(rRS) = True Then
		rRS.PageSize = vPageSize
		If IsNumber(vPage,0,False) = False Then vPage = 1
		If rRS.PageCount < CInt(vPage) Then vPage = rRS.PageCount

		If vSort = "" Then
			rRS.Sort = "SortNum,UnitCost"
		Else
			rRS.Sort = "SortNum DESC,UnitCost DESC"
		End If

		iCost = 0
		iAdoptNumPlan = 0
		iAdoptNum = 0
		Do While GetRSState(rRS) = True
			dbCost = rRS.Collect("Cost")
			dbAdoptNumPlan = rRS.Collect("AdoptNumPlan")
			dbAdoptNumResult = rRS.Collect("AdoptNumResult")
			dbAdoptNumPlanPeriod = rRS.Collect("AdoptNumPlanPeriod")

			iCost = iCost + dbCost
			iAdoptNumPlan = iAdoptNumPlan + dbAdoptNumPlanPeriod
			iAdoptNum = iAdoptNum + dbAdoptNumResult

			rRS.MoveNext
		Loop
		rRS.MoveFirst
		iUnitCost = 0
		If iAdoptNum > 0 Then iUnitCost = iCost / iAdoptNum

		If vBranchName <> "" Then
			sFilter = ""
			aBN = Split(Replace(vBranchName,"�@"," ")," ")
			For idx = 0 To UBound(aBN)
				If sFilter <> "" Then sFilter = sFilter & "OR "
				sFilter = sFilter & "BranchName LIKE '%" & aBN(idx) & "%' "
			Next
			rRS.Filter = Trim(sFilter)
		End If
	End If

	If GetRSState(rRS) = True Then
		sHTML = sHTML & "<table class=""pattern6"" border=""0"" style=""width:100%;margin-bottom:15px;"">"
		sHTML = sHTML & "<colgroup>"
		sHTML = sHTML & "<col style=""width:5%;"">"
		sHTML = sHTML & "<col style=""width:19%;"">"
		sHTML = sHTML & "<col style=""width:16%;"">"
		sHTML = sHTML & "<col style=""width:16%;"">"
		sHTML = sHTML & "<col style=""width:16%;"">"
		sHTML = sHTML & "<col style=""width:25%;"">"
		sHTML = sHTML & "</colgroup>"
		sHTML = sHTML & "<thead>"
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th></th>"
		sHTML = sHTML & "<th style=""text-align:center;"">����E�X��</th>"
		sHTML = sHTML & "<th style=""text-align:center;"">�̗p��đ��z</th>"
		sHTML = sHTML & "<th style=""text-align:center;"">�̗p�l��<br>(����/�v��)</th>"
		sHTML = sHTML & "<th style=""text-align:center;"">1���̗p���</th>"
		sHTML = sHTML & "<th style=""text-align:center;"">�}��</th>"
		sHTML = sHTML & "</tr>"
		sHTML = sHTML & "</thead>"
		sHTML = sHTML & "<tbody>"

		rRS.AbsolutePage = vPage
		idx = 0
		Do While GetRSState(rRS) = True And idx < vPageSize
			dbCompanyCode = rRS.Collect("CompanyCode")
			dbBranchSeq = rRS.Collect("BranchSeq")
			dbBranchName = rRS.Collect("BranchName")
			dbCost = rRS.Collect("Cost")
			dbAdoptNumPlan = rRS.Collect("AdoptNumPlan")
			dbAdoptNumResult = rRS.Collect("AdoptNumResult")
			dbUnitCost = rRS.Collect("UnitCost")
			dbMedName = rRS.Collect("MedName")
			dbAdoptNumPlanPeriod = rRS.Collect("AdoptNumPlanPeriod")

			sFlag = ""
			If dbUnitCost > iUnitCost Or dbUnitCost = 0 Then sFlag = "��"

			sHref = rCP.GetSearchParam()
			If sHref <> "" Then
				sHref = HTTPS_CURRENTURL & "company/costperformance/branch.asp" & sHref & "&dcc=" & dbCompanyCode & "&branchseq=" & dbBranchSeq
			Else
				sHref = HTTPS_CURRENTURL & "company/costperformance/branch.asp?dcc=" & dbCompanyCode & "&branchseq=" & dbBranchSeq
			End If

			sHTML = sHTML & "<tr>"
			sHTML = sHTML & "<td>" & sFlag & "</td>"
			sHTML = sHTML & "<td><a href=""" & Replace(sHref,"&","&amp;") & """>" & dbBranchName & "</a></td>"
			sHTML = sHTML & "<td style=""text-align:right;"">" & FormatCurrency(Round(dbCost)) & "</td>"
			sHTML = sHTML & "<td style=""text-align:right;"">"
			If dbAdoptNumResult > 0 Then
				sHTML = sHTML & dbAdoptNumResult
			Else
				sHTML = sHTML & "-&nbsp;"
			End If
			sHTML = sHTML & "/" & RoundUp(dbAdoptNumPlanPeriod,0) & "��"
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "<td style=""text-align:right;"">"
			If dbUnitCost > 0 Then
				sHTML = sHTML & FormatCurrency(Round(dbUnitCost)) & "/��"
			Else
				sHTML = sHTML & "-"
			End If
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "<td>" & dbMedName & "</td>"
			sHTML = sHTML & "</tr>"

			idx = idx + 1
			rRS.MoveNext
		Loop

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th colspan=""2"" style=""border-top:2px solid #cecfff;"">�S����E�X�܂̍��v�E����</th>"
		sHTML = sHTML & "<td style=""border-top:2px solid #cecfff;text-align:right;"">" & FormatCurrency(Round(iCost)) & "</td>"
		sHTML = sHTML & "<td style=""border-top:2px solid #cecfff;text-align:right;"">" & iAdoptNum & "/" & RoundUp(iAdoptNumPlan,0) & "��</td>"
		sHTML = sHTML & "<td style=""border-top:2px solid #cecfff;text-align:right;"">"
		If iUnitCost > 0 Then
			sHTML = sHTML & FormatCurrency(Round(iUnitCost)) & "/��"
		Else
			sHTML = sHTML & "-"
		End If
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<th style=""border-top:2px solid #cecfff;""></th>"
		sHTML = sHTML & "</tr>"

		sHTML = sHTML & "</tbody>"
		sHTML = sHTML & "</table>"

		rRS.MoveFirst
	End If

	htmlCostPerformance_Branch = sHTML
End Function
%>
