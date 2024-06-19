<%
'*******************************************************************************
'�T�@�v�F�̗p���P�T�|�[�g�V�X�e���̔}�̕ʔ�p�Ό���TABLE���擾
'���@���FrDB	�F�ڑ���DB�R�l�N�V����
'�@�@�@�FvUserID�F���O�C�������[�UID
'�@�@�@�FvYM1	�F�W�v���ԉ����N��
'�@�@�@�FvYM2	�F�W�v���ԏ���N��
'�߂�l�FString
'���@�l�F
'���@���F2010/03/10 LIS K.Kokubo �쐬
'*******************************************************************************
Function htmlCostPerformance_Media(ByRef rDB, ByRef rRS, ByVal vUserID, ByVal rCP, ByVal vPageSize, ByVal vPage, ByVal vSort)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sSQLErr
	'DB
	Dim dbMedName
	Dim dbAdoptNumPlan
	Dim dbAdoptNumResult
	Dim dbCost
	Dim dbUnitCost
	Dim dbUnitCostRank
	Dim dbAdoptNumResultRank
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
	End If

	If GetRSState(rRS) = True Then
		sHTML = sHTML & "<table class=""pattern6"" style=""width:100%;"">"
		sHTML = sHTML & "<colgroup>"
		sHTML = sHTML & "<col style=""width:8%;"">"
		sHTML = sHTML & "<col style=""width:36%;"">"
		sHTML = sHTML & "<col style=""width:20%;"">"
		sHTML = sHTML & "<col style=""width:16%;"">"
		sHTML = sHTML & "<col style=""width:20%;"">"
		sHTML = sHTML & "</colgroup>"
		sHTML = sHTML & "<thead>"
		sHTML = sHTML & "<th>����</th>"
		sHTML = sHTML & "<th>�}�̖�</th>"
		sHTML = sHTML & "<th>�̗p�R�X�g���z</th>"
		sHTML = sHTML & "<th>�̗p�l��<br>(����/�v��)</th>"
		sHTML = sHTML & "<th>�P���̗p�R�X�g</th>"
		sHTML = sHTML & "</thead>"
		sHTML = sHTML & "<tbody>"

		rRS.AbsolutePage = vPage
		idx = 0
		Do While GetRSState(rRS) And idx < vPageSize
			dbMedName = rRS.Collect("MedName")
			dbAdoptNumPlan = rRS.Collect("AdoptNumPlan")
			dbAdoptNumResult = rRS.Collect("AdoptNumResult")
			dbCost = rRS.Collect("Cost")
			dbUnitCost = rRS.Collect("UnitCost")
			dbUnitCostRank = rRS.Collect("UnitCostRank")
			dbAdoptNumResultRank = rRS.Collect("AdoptNumResultRank")
			dbAdoptNumPlanPeriod = rRS.Collect("AdoptNumPlanPeriod")

			sHref = Replace(rCP.GetSearchParam(),"&","&amp;")
			sHref = sHref & "&amp;mn=" & Server.URLEncode(dbMedName)
			sHref = HTTPS_CURRENTURL & "company/costperformance/media.asp" & sHref

			sHTML = sHTML & "<tr>"
			sHTML = sHTML & "<td>"
			If dbAdoptNumResult > 0 Then
				If vSort = "" Then
					sHTML = sHTML & dbUnitCostRank
				Else
					sHTML = sHTML & dbAdoptNumResultRank
				End If
				sHTML = sHTML & "��"
				If (vSort = "" And CInt(dbUnitCostRank) = 1) Or (vSort <> "" And CInt(dbAdoptNumResultRank) = 1) Then sHTML = sHTML & "<img src=""/img/staff/rank_item.gif"" alt="""">"
			Else
				sHTML = sHTML & "-"
			End If
			sHTML = sHTML & "</td>"
			sHTML = sHTML & "<td><a href=""" & sHref & """>" & dbMedName & "</a></td>"
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
			sHTML = sHTML & "</tr>"

			idx = idx + 1
			rRS.MoveNext
		Loop

		sHTML = sHTML & "</tbody>"
		sHTML = sHTML & "</table>"
	End If

	htmlCostPerformance_Media = sHTML
End Function
%>
