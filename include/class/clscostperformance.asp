<%
'******************************************************************************
'�T�@�v�F�̗p���P�T�|�[�g�V�X�e���̃N���X
'�ց@���F��Public
'�@�@�@�FsqlCompareFranchisee		�F���c�X�E�e�b�X�܂̎��т��r����SQL�𐶐�
'�@�@�@�FsqlSearchCostPerformance	�F���Г��󋵂̏����ɂ�镔��E�X�܃f�[�^���oSQL�𐶐�
'�@�@�@�FsqlCostPerformanceMedia	�F���Г��󋵂̏����ɂ��}�̕ʔ�p�Ό��ʃf�[�^���oSQL�𐶐�
'�@�@�@�FsqlBranchDetail			�F����E�X�܏ڍׂ̔}�̕ʎ���SQL�𐶐�
'�@�@�@�FsqlMediaDetail				�F�}�̏ڍׂ̕���E�X�ܕʎ���SQL�𐶐�
'�@�@�@�FsqlYearChange				�F�N�Ԑ��ڃf�[�^�擾SQL
'�@�@�@�FsqlYearBranch				�F�N�Ԑ��ڂ̕���E�X�܈ꗗ
'�@�@�@�FsqlYearChangeBranch		�F����E�X�ܖ��̔N�Ԑ��ڃf�[�^���擾����SQL
'�@�@�@�FsqlSimulation				�F�V�~�����[�V�����f�t�H���g�f�[�^���oSQL�𐶐�
'�@�@�@�FsqlSimReferenceData		�F�V�~�����[�V�����Q�l�f�[�^
'�@�@�@�FsqlMedCompare				�F�����}�̔�rSQL�𐶐�
'�@�@�@�FsqlMedCompareYear1			�F�}�̔�r�N�Ԑ���SQL�𐶐�
'�@�@�@�FsqlMedCompareYear2			�F�}�̔�r�N�Ԑ���SQL�𐶐�
'�@�@�@�F��Private
'�@�@�@�F
'���@�l�F������ �ڍ׌����p�p�����[�^ �i�A�h�z�b�N�Ȃr�p�k�����j
'���@���F2009/10/23 LIS K.Kokubo �쐬
'******************************************************************************
Class clsCostPerformance
	Public CompanyCode
	'��������
	Public DspCompanyCode
	Public BranchSeq
	Public Y1
	Public Y2
	Public M1
	Public M2
	Public YM1
	Public YM2
	Public WorkingTypeCode
	Public JobTypeCode
	Public PrefectureCode
	Public IndustryTypeCode
	Public BranchName
	Public MedName

	'******************************************************************************
	'�T�@�v�F�R���X�g���N�^
	'���@���F
	'���@�l�F
	'���@���F2009/10/23 LIS K.Kokubo �쐬
	'******************************************************************************
	Private Sub Class_Initialize()
		CompanyCode = G_USERID

		'�p�����[�^���猟���������擾
		Call ReadParam()
	End Sub

	'******************************************************************************
	'�T�@�v�FGET�f�[�^�̓ǂݍ���
	'���@���F
	'���@�l�F
	'���@���F2009/10/23 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Sub ReadParam()
		Y1 = GetForm("y1",2)
		Y2 = GetForm("y2",2)
		M1 = GetForm("m1",2)
		M2 = GetForm("m2",2)
		If Y1 <> "" And M1 <> "" Then YM1 = Y1 & Right("0" & M1,2)
		If Y2 <> "" And M2 <> "" Then YM2 = Y2 & Right("0" & M2,2)

		If GetForm("dcc",2) <> "" Then DspCompanyCode = GetForm("dcc",2)
		If GetForm("branchseq",2) <> "" Then BranchSeq = GetForm("branchseq",2)
		If GetForm("wt",2) <> "" Then WorkingTypeCode = GetForm("wt",2)
		If GetForm("jt",2) <> "" Then JobTypeCode = GetForm("jt",2)
		If GetForm("pc",2) <> "" Then PrefectureCode = GetForm("pc",2)
		If GetForm("it",2) <> "" Then IndustryTypeCode = GetForm("it",2)
		If GetForm("bn",2) <> "" Then BranchName = GetForm("bn",2)
		If GetForm("mn",2) <> "" Then MedName = GetForm("mn",2)

		'�f�[�^�������`�F�b�N
		Call ChkData()

		'�R�[�h�Ή����̎擾
		Call SetData()
	End Sub

	'******************************************************************************
	'�T�@�v�F�R�[�h�ɑΉ��������̂��擾����
	'���@���F
	'���@�l�F
	'���@���F2009/10/23 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Sub SetData()
	End Sub

	'******************************************************************************
	'�T�@�v�F�f�[�^�̐��������`�F�b�N
	'���@���F
	'���@�l�F
	'���@���F2009/10/23 LIS K.Kokubo �쐬
	'******************************************************************************
	Private Sub ChkData()
		If IsRE(YM1,"^20\d\d((0[123456789])|(1[012]))$",True) = False Then YM1 = ""
		If IsRE(YM2,"^20\d\d((0[123456789])|(1[012]))$",True) = False Then YM2 = ""
		If YM1 & YM2 = "" Then
			If IsRE(YM1,"^20\d\d((0[123456789])|(1[012]))$",True) = False Then YM1 = Year(Date) & Right(("0" & Month(Date)),2)
			If IsRE(YM2,"^20\d\d((0[123456789])|(1[012]))$",True) = False Then YM2 = Year(Date) & Right(("0" & Month(Date)),2)
		End If
	End Sub

	'******************************************************************************
	'�T�@�v�F���d���ڍ׌����y�[�W�֓n��GET�p�����[�^�𐶐����Ď擾�B
	'���@���F
	'���@�l�F
	'���@���F2009/10/23 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Function GetSearchParam()
		Dim idx
		Dim sParam

		sParam = ""

		If YM1 <> "" Then
			sParam = sParam & "&y1=" & Y1
			sParam = sParam & "&m1=" & M1
		End If
		If YM2 <> "" Then
			sParam = sParam & "&y2=" & Y2
			sParam = sParam & "&m2=" & M2
		End If
		If DspCompanyCode <> "" Then sParam = sParam & "&dcc=" & DspCompanyCode
		If WorkingTypeCode <> "" Then sParam = sParam & "&wt=" & WorkingTypeCode
		If JobTypeCode <> "" Then sParam = sParam & "&jt=" & JobTypeCode
		If PrefectureCode <> "" Then sParam = sParam & "&pc=" & PrefectureCode
		If IndustryTypeCode <> "" Then sParam = sParam & "&it=" & IndustryTypeCode
		If BranchName <> "" Then sParam = sParam & "&bn=" & Server.URLEncode(BranchName)

		If sParam <> "" Then
			'����&���H�ɕϊ�
			sParam = "?" & Mid(sParam, 2)

			'�h�d�̎d�l�̓p�����[�^�̏�����Q�O�S�W�o�C�g
			sParam = Left(sParam, 2048)
		End If
		GetSearchParam = sParam
	End Function

	'******************************************************************************
	'�T�@�v�FSQL�쐬�ɗ��p����f�[�^���擾
	'���@���FrDeclare	�F[OUTPUT]sp_executesql��@params
	'�@�@�@�FrParams	�F[OUTPUT]sp_executesql��@param1...
	'�@�@�@�FrJoin		�F[OUTPUT]JOIN
	'�@�@�@�FrWhere		�F[OUTPUT]WHERE
	'�@�@�@�FvCSV		�F���p�f�[�^�w�� ��..."CompanyCode,BranchSeq"
	'���@�l�F
	'���@���F2009/10/23 LIS K.Kokubo �쐬
	'******************************************************************************
	Function setSQLData(ByRef rDeclare, ByRef rParams, ByRef rJoin, ByRef rWhere, ByVal vCSV)
		Dim tmp1
		Dim tmp2
		Dim iPrmNo
		Dim aValue
		Dim idx
		Dim aCSV

		If ChkStr(vCSV) = "" Then Exit Function

		aCSV = Split(Replace(vCSV," ",""),",")

		'<��ƃR�[�h>
		If CompanyCode <> "" And UBound(Filter(aCSV,"CompanyCode")) >= 0 Then
			If rDeclare <> "" Then rDeclare = rDeclare & ","
			rDeclare = rDeclare & "@vCompanyCode VARCHAR(8)"
			rParams = rParams & ",@vCompanyCode = N'" & CompanyCode & "'"

			If rWhere <> "" Then rWhere = rWhere & "AND "
			'rWhere = rWhere & "A.CompanyCode = @vCompanyCode "
			rWhere = rWhere & "EXISTS(SELECT * FROM (SELECT @vCompanyCode AS CompanyCode UNION SELECT CompanyCode FROM CompanyInfo WHERE GroupCode = @vCompanyCode) AS TMP WHERE A.CompanyCode = TMP.CompanyCode) "
		End If
		'</��ƃR�[�h>

		'<�\���Ώۊ�ƃR�[�h>
		If CompanyCode <> "" And UBound(Filter(aCSV,"DspCompanyCode")) >= 0 Then
			If rDeclare <> "" Then rDeclare = rDeclare & ","
			rDeclare = rDeclare & "@vDspCompanyCode VARCHAR(8)"
			rParams = rParams & ",@vDspCompanyCode = N'" & DspCompanyCode & "'"

			If rWhere <> "" Then rWhere = rWhere & "AND "
			rWhere = rWhere & "A.CompanyCode = @vDspCompanyCode "
		End If
		'</�\���Ώۊ�ƃR�[�h>

		'<����E�X�ܔԍ�>
		If BranchSeq <> "" And UBound(Filter(aCSV,"BranchSeq")) >= 0 Then
			If rDeclare <> "" Then rDeclare = rDeclare & ","
			rDeclare = rDeclare & "@vBranchSeq INT"
			rParams = rParams & ",@vBranchSeq = N'" & BranchSeq & "'"

			If rWhere <> "" Then rWhere = rWhere & "AND "
			rWhere = rWhere & "A.BranchSeq = @vBranchSeq "
		End If
		'</����E�X�ܔԍ�>

		'<�W�v���ԉ���>
		If YM1 <> "" And UBound(Filter(aCSV,"StartDay")) >= 0 Then
			If rDeclare <> "" Then rDeclare = rDeclare & ","
			rDeclare = rDeclare & "@vStartYM VARCHAR(6)"
			rParams = rParams & ",@vStartYM = N'" & YM1 & "'"

			If rDeclare <> "" Then rDeclare = rDeclare & ","
			rDeclare = rDeclare & "@dStartDay DATETIME"
			rParams = rParams & ",@dStartDay = N'" & GetDateStr(CDate(Left(YM1,4) & "/" & Right(YM1,2) & "/01"),"") & "'"
		End If
		'</�W�v���ԉ���>

		'<�W�v���ԏ��>
		If YM2 <> "" And UBound(Filter(aCSV,"EndDay")) >= 0 Then
			If rDeclare <> "" Then rDeclare = rDeclare & ","
			rDeclare = rDeclare & "@vEndYM VARCHAR(6)"
			rParams = rParams & ",@vEndYM = N'" & YM2 & "'"

			If rDeclare <> "" Then rDeclare = rDeclare & ","
			rDeclare = rDeclare & "@dEndDay DATETIME"
			rParams = rParams & ",@dEndDay = N'" & GetDateStr(DateAdd("d",-1,DateAdd("m",1,CDate(Left(YM2,4) & "/" & Right(YM2,2) & "/01"))),"") & "'"
		End If
		'</�W�v���ԏ��>

		'<�Ζ��`��>
		tmp1 = ""
		iPrmNo = 1
		If WorkingTypeCode <> "" And UBound(Filter(aCSV,"WorkingTypeCode")) >= 0 Then
			aValue = Split(WorkingTypeCode,",")
			For idx = 0 To UBound(aValue)
				If aValue(idx) <> "" Then
					If rDeclare <> "" Then rDeclare = rDeclare & ","
					rDeclare = rDeclare & "@vWorkingTypeCode" & iPrmNo & " VARCHAR(3)"
					rParams = rParams & ",@vWorkingTypeCode" & iPrmNo & " = N'" & aValue(idx) & "'"

					If tmp1 <> "" Then tmp1 = tmp1 & ","
					tmp1 = tmp1 & "@vWorkingTypeCode" & iPrmNo

					iPrmNo = iPrmNo + 1
				End If
			Next

			If rWhere <> "" Then rWhere = rWhere & "AND "
			rWhere = rWhere & "A.WorkingTypeCode IN (" & tmp1 & ") "
		End If
		'<�Ζ��`��>

		'<�E��>
		tmp1 = ""
		iPrmNo = 1
		If JobTypeCode <> "" And UBound(Filter(aCSV,"JobTypeCode")) >= 0 Then
			aValue = Split(JobTypeCode,",")
			For idx = 0 To UBound(aValue)
				If aValue(idx) <> "" Then
					If rDeclare <> "" Then rDeclare = rDeclare & ","
					rDeclare = rDeclare & "@vJobTypeCode" & iPrmNo & " VARCHAR(7)"
					rParams = rParams & ",@vJobTypeCode" & iPrmNo & " = N'" & aValue(idx) & "'"

					If tmp1 <> "" Then tmp1 = tmp1 & ","
					tmp1 = tmp1 & "@vJobTypeCode" & iPrmNo

					iPrmNo = iPrmNo + 1
				End If
			Next

			If rWhere <> "" Then rWhere = rWhere & "AND "
			rWhere = rWhere & "A.JobTypeCode IN (" & tmp1 & ") "
		End If
		'</�E��>

		'<�s���{��>
		tmp1 = ""
		iPrmNo = 1
		If PrefectureCode <> "" And UBound(Filter(aCSV,"PrefectureCode")) >= 0 Then
			aValue = Split(PrefectureCode,",")
			For idx = 0 To UBound(aValue)
				If aValue(idx) <> "" Then
					If rDeclare <> "" Then rDeclare = rDeclare & ","
					rDeclare = rDeclare & "@vPrefectureCode" & iPrmNo & " VARCHAR(3)"
					rParams = rParams & ",@vPrefectureCode" & iPrmNo & " = N'" & aValue(idx) & "'"

					If tmp1 <> "" Then tmp1 = tmp1 & ","
					tmp1 = tmp1 & "@vPrefectureCode" & iPrmNo

					iPrmNo = iPrmNo + 1
				End If
			Next

			If rWhere <> "" Then rWhere = rWhere & "AND "
			rWhere = rWhere & "A.PrefectureCode IN (" & tmp1 & ") "
		End If
		'</�s���{��>

		'<�Ǝ�>
		tmp1 = ""
		iPrmNo = 1
		If IndustryTypeCode <> "" And UBound(Filter(aCSV,"IndustryTypeCode")) >= 0 Then
			aValue = Split(IndustryTypeCode,",")
			For idx = 0 To UBound(aValue)
				If aValue(idx) <> "" Then
					If rDeclare <> "" Then rDeclare = rDeclare & ","
					rDeclare = rDeclare & "@vIndustryTypeCode" & iPrmNo & " VARCHAR(3)"
					rParams = rParams & ",@vIndustryTypeCode" & iPrmNo & " = N'" & aValue(idx) & "'"

					If tmp1 <> "" Then tmp1 = tmp1 & ","
					tmp1 = tmp1 & "@vIndustryTypeCode" & iPrmNo

					iPrmNo = iPrmNo + 1
				End If
			Next

'			If rWhere <> "" Then rWhere = rWhere & "AND "
'			rWhere = rWhere & "A.IndustryTypeCode IN (" & tmp1 & ") "
		End If
		'</�Ǝ�>

		'<�}�̖�>
		If MedName <> "" And UBound(Filter(aCSV,"MedName")) >= 0 Then
			If rDeclare <> "" Then rDeclare = rDeclare & ","
			rDeclare = rDeclare & "@vMedName VARCHAR(100)"
			rParams = rParams & ",@vMedName = N'" & MedName & "'"

			If rWhere <> "" Then rWhere = rWhere & "AND "
			rWhere = rWhere & "A.MedName = @vMedName "
		End If
		'</�}�̖�>
	End Function

	'******************************************************************************
	'�T�@�v�F���c�X�E�e�b�X�܂̎��т��r����SQL�𐶐�
	'���@���F
	'���@�l�F
	'���@���F2010/03/10 LIS K.Kokubo �쐬
	'******************************************************************************
	Function sqlCompareFranchisee()
		Dim sDeclare
		Dim sParams
		Dim sJoin
		Dim sWhere

		Dim sSQL
		Dim sSQL2

		sDeclare = ""
		sParams = ""
		sWhere = ""

		'�f�[�^�������`�F�b�N
		Call ChkData()

		Call setSQLData(sDeclare,sParams,sJoin,sWhere,"CompanyCode,BranchSeq,StartDay,EndDay,WorkingTypeCode,JobTypeCode,PrefectureCode,IndustryTypeCode,MedName")

		If sWhere <> "" Then
			sWhere = "WHERE " & sWhere
		End If

		sSQL = ""
		sSQL = sSQL & "SELECT A.FranchiseeFlag,A.Cost,A.AdoptNumPlan,A.AdoptNumResult,A.UnitCost,A.AdoptNumPlanPeriod,CASE WHEN A.AdoptNumResult > 0 THEN 1 ELSE 2 END AS SortNum,RANK() OVER(ORDER BY CASE WHEN A.AdoptNumResult > 0 THEN 1 ELSE 2 END,A.UnitCost ASC) AS UnitCostRank,RANK() OVER(ORDER BY A.AdoptNumResult DESC) AS AdoptNumResultRank FROM (SELECT A.FranchiseeFlag,SUM(A.GroupCost) AS Cost,SUM(A.AdoptNumPlan) AS AdoptNumPlan,SUM(A.AdoptNumResult) AS AdoptNumResult,CASE WHEN SUM(A.AdoptNumResult) > 0 THEN SUM(A.GroupCost) / SUM(A.AdoptNumResult) ELSE 0 END AS UnitCost,SUM(A.AdoptNumPlanPeriod) AS AdoptNumPlanPeriod FROM (SELECT A.FranchiseeFlag,A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,C.WorkingTypeCode,C.JobTypeCode,C.PrefectureCode,C.AdoptNum AS AdoptNumPlan,COALESCE(D.AdoptNum,0) AS AdoptNumResult,CASE WHEN COALESCE(B.AdoptNum,0) > 0 THEN A.Cost * CONVERT(FLOAT,C.AdoptNum) / CONVERT(FLOAT,B.AdoptNum) ELSE 0 END AS GroupCost,CONVERT(FLOAT,B.AdoptNum)*Period AS AdoptNumPlanPeriod FROM (SELECT A.FranchiseeFlag,A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.Cost * (CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1)) AS Cost,CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1) AS Period FROM (SELECT B.FranchiseeFlag,A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.EndDay,CONVERT(FLOAT,A.Cost) AS Cost,CASE WHEN A.StartDay >= @dStartDay THEN A.StartDay ELSE @dStartDay END AS StartDay2,CASE WHEN A.EndDay <= @dEndDay THEN A.EndDay ELSE @dEndDay END AS EndDay2 FROM CMPCostPerformanceMedia AS A INNER JOIN CMPCostPerformanceBranch AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq WHERE (@vStartYM BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (@vEndYM BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (@vStartYM <= CONVERT(VARCHAR(6),A.StartDay,112) AND @vEndYM >= CONVERT(VARCHAR(6),A.EndDay,112))) AS A) AS A INNER JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptPlan AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay) AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay INNER JOIN CMPCostPerformanceAdoptPlan AS C ON A.CompanyCode = C.CompanyCode AND A.BranchSeq = C.BranchSeq AND A.MedName = C.MedName AND A.StartDay = C.StartDay LEFT JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE (A.AdoptMonth BETWEEN @vStartYM AND @vEndYM) AND A.AdoptMonth < CONVERT(VARCHAR(6),B.EndDay,112) GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode UNION ALL SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE A.AdoptMonth >= CONVERT(VARCHAR(6),B.EndDay,112) AND CONVERT(VARCHAR(6),B.EndDay,112) <= @vEndYM GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS D ON C.CompanyCode = D.CompanyCode AND C.BranchSeq = D.BranchSeq AND C.MedName = D.MedName AND C.StartDay = D.StartDay AND C.JobTypeCode = D.JobTypeCode AND C.WorkingTypeCode = D.WorkingTypeCode AND C.PrefectureCode = D.PrefectureCode) AS A "
		sSQL = sSQL & sWhere
		sSQL = sSQL & "GROUP BY A.FranchiseeFlag) AS A ORDER BY SortNum,UnitCost;"

		sSQL2 = ""
		sSQL2 = sSQL2 & "/* �i�r�E�̗p���P�T�|�[�g�V�X�e�� ���c�X�E�e�b�X�� */" & vbCrLf
		sSQL2 = sSQL2 & "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED" & vbCrLf
		sSQL2 = sSQL2 & "EXEC sp_executesql N'" & Replace(sSQL, "'", "''") & "'"
		If sDeclare <> "" Then sSQL2 = sSQL2 & vbCrLf & ",N'" & sDeclare & "'" & vbCrLf & sParams

		sqlCompareFranchisee = sSQL2 & vbCrLf
	End Function

	'******************************************************************************
	'�T�@�v�F���Г��󋵂̏����ɂ�镔��E�X�܃f�[�^���oSQL�𐶐�
	'���@���F
	'���@�l�F
	'���@���F2009/10/23 LIS K.Kokubo �쐬
	'******************************************************************************
	Function sqlSearchCostPerformance()
		Dim sDeclare
		Dim sParams
		Dim sJoin
		Dim sWhere
		Dim tmpParam

		Dim sSQL
		Dim sSQL2

		sDeclare = ""
		sParams = ""
		sWhere = ""

		'�f�[�^�������`�F�b�N
		Call ChkData()

		Call setSQLData(sDeclare,sParams,sJoin,sWhere,"CompanyCode,BranchSeq,StartDay,EndDay,WorkingTypeCode,JobTypeCode,PrefectureCode,IndustryTypeCode,MedName")

		If sWhere <> "" Then
			sWhere = "WHERE " & sWhere
		End If

		'dbo.ftbl_CMPCostPerformanceBranch_Media(@vCompanyCode,@vStartYM,@vEndYM," & tmpParam & ",',')
		'XML�`���ɂ��Ȃ��Ƃ܂��������B
		tmpParam = ""
		If WorkingTypeCode <> "" Then: tmpParam = tmpParam & "@vWorkingTypeCode1": Else: tmpParam = tmpParam & "''": End If: tmpParam = tmpParam & ","
		If PrefectureCode <> "" Then: tmpParam = tmpParam & "@vPrefectureCode1": Else: tmpParam = tmpParam & "''": End If: tmpParam = tmpParam & ","
		If JobTypeCode <> "" Then: tmpParam = tmpParam & "@vJobTypeCode1": Else: tmpParam = tmpParam & "''": End If: tmpParam = tmpParam & ","
		If IndustryTypeCode <> "" Then: tmpParam = tmpParam & "@vIndustryTypeCode1": Else: tmpParam = tmpParam & "''": End If

		sSQL = ""
		sSQL = sSQL & "SELECT CASE WHEN UnitCost > 0 THEN 1 ELSE 2 END AS SortNum,A.CompanyCode,A.BranchSeq,A.Cost,A.AdoptNumPlan,A.AdoptNumResult,A.UnitCost,B.BranchName,C.MedName,A.AdoptNumPlanPeriod FROM (SELECT A.CompanyCode,A.BranchSeq,SUM(A.GroupCost) AS Cost,SUM(A.AdoptNumPlan) AS AdoptNumPlan,SUM(A.AdoptNumResult) AS AdoptNumResult,CASE WHEN SUM(A.AdoptNumResult) > 0 THEN SUM(A.GroupCost) / SUM(A.AdoptNumResult) ELSE 0 END AS UnitCost,SUM(A.AdoptNumPlanPeriod) AS AdoptNumPlanPeriod FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,C.WorkingTypeCode,C.JobTypeCode,C.PrefectureCode,C.AdoptNum AS AdoptNumPlan,COALESCE(D.AdoptNum,0) AS AdoptNumResult,CASE WHEN COALESCE(B.AdoptNum,0) > 0 THEN A.Cost * CONVERT(FLOAT,C.AdoptNum) / CONVERT(FLOAT,B.AdoptNum) ELSE 0 END AS GroupCost,CONVERT(FLOAT,B.AdoptNum)*Period AS AdoptNumPlanPeriod FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.Cost * (CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1)) AS Cost,CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1) AS Period FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.EndDay,CONVERT(FLOAT,A.Cost) AS Cost,CASE WHEN A.StartDay >= @dStartDay THEN A.StartDay ELSE @dStartDay END AS StartDay2,CASE WHEN A.EndDay <= @dEndDay THEN A.EndDay ELSE @dEndDay END AS EndDay2 FROM CMPCostPerformanceMedia AS A WHERE (@vStartYM BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (@vEndYM BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (@vStartYM <= CONVERT(VARCHAR(6),A.StartDay,112) AND @vEndYM >= CONVERT(VARCHAR(6),A.EndDay,112))) AS A) AS A INNER JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptPlan AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay) AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay INNER JOIN CMPCostPerformanceAdoptPlan AS C ON A.CompanyCode = C.CompanyCode AND A.BranchSeq = C.BranchSeq AND A.MedName = C.MedName AND A.StartDay = C.StartDay LEFT JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE (A.AdoptMonth BETWEEN @vStartYM AND @vEndYM) AND A.AdoptMonth < CONVERT(VARCHAR(6),B.EndDay,112) GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode UNION ALL SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE A.AdoptMonth >= CONVERT(VARCHAR(6),B.EndDay,112) AND CONVERT(VARCHAR(6),B.EndDay,112) <= @vEndYM GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS D ON C.CompanyCode = D.CompanyCode AND C.BranchSeq = D.BranchSeq AND C.MedName = D.MedName AND C.StartDay = D.StartDay AND C.JobTypeCode = D.JobTypeCode AND C.WorkingTypeCode = D.WorkingTypeCode AND C.PrefectureCode = D.PrefectureCode) AS A "
		sSQL = sSQL & sWhere
		sSQL = sSQL & "GROUP BY A.CompanyCode,A.BranchSeq) AS A INNER JOIN CMPCostPerformanceBranch AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq INNER JOIN dbo.ftbl_CMPCostPerformanceBranch_Media(@vCompanyCode,@vStartYM,@vEndYM," & tmpParam & ",',') AS C ON A.CompanyCode = C.CompanyCode AND A.BranchSeq = C.BranchSeq;"

		sSQL2 = ""
		sSQL2 = sSQL2 & "/* �i�r�E�̗p���P�T�|�[�g�V�X�e�� ���Г��󋵕���E�X�� */" & vbCrLf
		sSQL2 = sSQL2 & "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED" & vbCrLf
		sSQL2 = sSQL2 & "EXEC sp_executesql N'" & Replace(sSQL, "'", "''") & "'"
		If sDeclare <> "" Then sSQL2 = sSQL2 & vbCrLf & ",N'" & sDeclare & "'" & vbCrLf & sParams

		sqlSearchCostPerformance = sSQL2 & vbCrLf
	End Function

	'******************************************************************************
	'�T�@�v�F���Г��󋵂̏����ɂ��}�̕ʔ�p�Ό��ʃf�[�^���oSQL�𐶐�
	'���@���F
	'���@�l�F
	'���@���F2010/01/20 LIS K.Kokubo �쐬
	'******************************************************************************
	Function sqlCostPerformanceMedia()
		Dim sDeclare
		Dim sParams
		Dim sJoin
		Dim sWhere

		Dim sSQL
		Dim sSQL2

		sDeclare = ""
		sParams = ""
		sWhere = ""

		'�f�[�^�������`�F�b�N
		Call ChkData()

		Call setSQLData(sDeclare,sParams,sJoin,sWhere,"CompanyCode,BranchSeq,StartDay,EndDay,WorkingTypeCode,JobTypeCode,PrefectureCode,IndustryTypeCode,MedName")

		If sWhere <> "" Then
			sWhere = "WHERE " & sWhere
		End If

		sSQL = ""
		sSQL = sSQL & "SELECT A.MedName,A.Cost,A.AdoptNumPlan,A.AdoptNumResult,A.UnitCost,A.AdoptNumPlanPeriod,CASE WHEN A.AdoptNumResult > 0 THEN 1 ELSE 2 END AS SortNum,RANK() OVER(ORDER BY CASE WHEN A.AdoptNumResult > 0 THEN 1 ELSE 2 END,A.UnitCost ASC) AS UnitCostRank,RANK() OVER(ORDER BY A.AdoptNumResult DESC) AS AdoptNumResultRank FROM (SELECT A.MedName,SUM(A.GroupCost) AS Cost,SUM(A.AdoptNumPlan) AS AdoptNumPlan,SUM(A.AdoptNumResult) AS AdoptNumResult,CASE WHEN SUM(A.AdoptNumResult) > 0 THEN SUM(A.GroupCost) / SUM(A.AdoptNumResult) ELSE 0 END AS UnitCost,SUM(A.AdoptNumPlanPeriod) AS AdoptNumPlanPeriod FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,C.WorkingTypeCode,C.JobTypeCode,C.PrefectureCode,C.AdoptNum AS AdoptNumPlan,COALESCE(D.AdoptNum,0) AS AdoptNumResult,CASE WHEN COALESCE(B.AdoptNum,0) > 0 THEN A.Cost * CONVERT(FLOAT,C.AdoptNum) / CONVERT(FLOAT,B.AdoptNum) ELSE 0 END AS GroupCost,CONVERT(FLOAT,B.AdoptNum)*Period AS AdoptNumPlanPeriod FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.Cost * (CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1)) AS Cost,CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1) AS Period FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.EndDay,CONVERT(FLOAT,A.Cost) AS Cost,CASE WHEN A.StartDay >= @dStartDay THEN A.StartDay ELSE @dStartDay END AS StartDay2,CASE WHEN A.EndDay <= @dEndDay THEN A.EndDay ELSE @dEndDay END AS EndDay2 FROM CMPCostPerformanceMedia AS A WHERE (@vStartYM BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (@vEndYM BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (@vStartYM <= CONVERT(VARCHAR(6),A.StartDay,112) AND @vEndYM >= CONVERT(VARCHAR(6),A.EndDay,112))) AS A) AS A INNER JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptPlan AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay) AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay INNER JOIN CMPCostPerformanceAdoptPlan AS C ON A.CompanyCode = C.CompanyCode AND A.BranchSeq = C.BranchSeq AND A.MedName = C.MedName AND A.StartDay = C.StartDay LEFT JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE (A.AdoptMonth BETWEEN @vStartYM AND @vEndYM) AND A.AdoptMonth < CONVERT(VARCHAR(6),B.EndDay,112) GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode UNION ALL SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE A.AdoptMonth >= CONVERT(VARCHAR(6),B.EndDay,112) AND CONVERT(VARCHAR(6),B.EndDay,112) <= @vEndYM GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS D ON C.CompanyCode = D.CompanyCode AND C.BranchSeq = D.BranchSeq AND C.MedName = D.MedName AND C.StartDay = D.StartDay AND C.JobTypeCode = D.JobTypeCode AND C.WorkingTypeCode = D.WorkingTypeCode AND C.PrefectureCode = D.PrefectureCode) AS A "
		sSQL = sSQL & sWhere
		sSQL = sSQL & "GROUP BY A.MedName) AS A ORDER BY SortNum,UnitCost;"

		sSQL2 = ""
		sSQL2 = sSQL2 & "/* �i�r�E�̗p���P�T�|�[�g�V�X�e�� ���Г��󋵔}�̕ʔ�p�Ό��ʈꗗ */" & vbCrLf
		sSQL2 = sSQL2 & "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED" & vbCrLf
		sSQL2 = sSQL2 & "EXEC sp_executesql N'" & Replace(sSQL, "'", "''") & "'"
		If sDeclare <> "" Then sSQL2 = sSQL2 & vbCrLf & ",N'" & sDeclare & "'" & vbCrLf & sParams

		sqlCostPerformanceMedia = sSQL2 & vbCrLf
	End Function

	'******************************************************************************
	'�T�@�v�F����E�X�܏ڍׂ̔}�̕ʎ���SQL�𐶐�
	'���@���F
	'���@�l�F
	'���@���F2009/10/23 LIS K.Kokubo �쐬
	'******************************************************************************
	Function sqlBranchDetail()
		Dim sDeclare
		Dim sParams
		Dim sJoin
		Dim sWhere

		Dim sSQL
		Dim sSQL2

		sDeclare = ""
		sParams = ""
		sWhere = ""

		'�f�[�^�������`�F�b�N
		Call ChkData()

		Call setSQLData(sDeclare,sParams,sJoin,sWhere,"CompanyCode,DspCompanyCode,BranchSeq,StartDay,EndDay,WorkingTypeCode,JobTypeCode,PrefectureCode,IndustryTypeCode,MedName")

		If sWhere <> "" Then
			sWhere = "WHERE " & sWhere
		End If

		sSQL = ""
		sSQL = sSQL & "SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.Cost,A.StartDay,A.AdoptNumPlan,A.AdoptNumResult,A.UnitCost,B.BranchName,C.WorkingTypeCode,C.WorkingTypeName,D.MiddleClassName AS JobTypeName,A.AdoptNumPlanPeriod FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.WorkingTypeCode,A.JobTypeCode,A.PrefectureCode,A.GroupCost AS Cost,A.AdoptNumPlan,A.AdoptNumResult,CASE WHEN A.AdoptNumResult > 0 THEN A.GroupCost / A.AdoptNumResult ELSE 0 END AS UnitCost,A.AdoptNumPlanPeriod FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,C.WorkingTypeCode,C.JobTypeCode,C.PrefectureCode,C.AdoptNum AS AdoptNumPlan,COALESCE(D.AdoptNum,0) AS AdoptNumResult,CASE WHEN COALESCE(B.AdoptNum,0) > 0 THEN A.Cost * CONVERT(FLOAT,C.AdoptNum) / CONVERT(FLOAT,B.AdoptNum) ELSE 0 END AS GroupCost,CONVERT(FLOAT,B.AdoptNum)*Period AS AdoptNumPlanPeriod FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.Cost * (CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1)) AS Cost,CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1) AS Period FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.EndDay,CONVERT(FLOAT,A.Cost) AS Cost,CASE WHEN A.StartDay >= @dStartDay THEN A.StartDay ELSE @dStartDay END AS StartDay2,CASE WHEN A.EndDay <= @dEndDay THEN A.EndDay ELSE @dEndDay END AS EndDay2 FROM CMPCostPerformanceMedia AS A WHERE (@vStartYM BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (@vEndYM BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (@vStartYM <= CONVERT(VARCHAR(6),A.StartDay,112) AND @vEndYM >= CONVERT(VARCHAR(6),A.EndDay,112))) AS A) AS A INNER JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptPlan AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay) AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay INNER JOIN CMPCostPerformanceAdoptPlan AS C ON A.CompanyCode = C.CompanyCode AND A.BranchSeq = C.BranchSeq AND A.MedName = C.MedName AND A.StartDay = C.StartDay LEFT JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE (A.AdoptMonth BETWEEN @vStartYM AND @vEndYM) AND A.AdoptMonth < CONVERT(VARCHAR(6),B.EndDay,112) GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode UNION ALL SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE A.AdoptMonth >= CONVERT(VARCHAR(6),B.EndDay,112) AND CONVERT(VARCHAR(6),B.EndDay,112) <= @vEndYM GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS D ON C.CompanyCode = D.CompanyCode AND C.BranchSeq = D.BranchSeq AND C.MedName = D.MedName AND C.StartDay = D.StartDay AND C.JobTypeCode = D.JobTypeCode AND C.WorkingTypeCode = D.WorkingTypeCode AND C.PrefectureCode = D.PrefectureCode) AS A "
		sSQL = sSQL & sWhere
		sSQL = sSQL & ") AS A INNER JOIN CMPCostPerformanceBranch AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq INNER JOIN vw_WorkingType AS C ON A.WorkingTypeCode = C.WorkingTypeCode INNER JOIN vw_JobType AS D ON A.JobTypeCode = D.JobTypeCode;"

		sSQL2 = ""
		sSQL2 = sSQL2 & "/* ����E�X�܏ڍׂ̔}�̕ʎ��� */" & vbCrLf
		sSQL2 = sSQL2 & "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED" & vbCrLf
		sSQL2 = sSQL2 & "EXEC sp_executesql N'" & Replace(sSQL, "'", "''") & "'"
		If sDeclare <> "" Then sSQL2 = sSQL2 & vbCrLf & ",N'" & sDeclare & "'" & vbCrLf & sParams

		sqlBranchDetail = sSQL2 & vbCrLf
	End Function

	'******************************************************************************
	'�T�@�v�F�}�̏ڍׂ̕���E�X�ܕʎ���SQL�𐶐�
	'���@���F
	'���@�l�F
	'���@���F2010/01/26 LIS K.Kokubo �쐬
	'******************************************************************************
	Function sqlMediaDetail()
		Dim sDeclare
		Dim sParams
		Dim sJoin
		Dim sWhere

		Dim sSQL
		Dim sSQL2

		sDeclare = ""
		sParams = ""
		sWhere = ""

		'�f�[�^�������`�F�b�N
		Call ChkData()

		Call setSQLData(sDeclare,sParams,sJoin,sWhere,"CompanyCode,BranchSeq,StartDay,EndDay,WorkingTypeCode,JobTypeCode,PrefectureCode,IndustryTypeCode,MedName")

		If sWhere <> "" Then
			sWhere = "WHERE " & sWhere
		End If

		sSQL = ""
		sSQL = sSQL & "SELECT RANK() OVER(ORDER BY CASE WHEN A.AdoptNumResult = 0 THEN 1 ELSE 0 END) AS SortNum,A.CompanyCode,A.BranchSeq,A.MedName,A.Cost,A.StartDay,A.AdoptNumPlan,A.AdoptNumResult,A.UnitCost,B.BranchName,A.AdoptNumPlanPeriod FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,SUM(A.GroupCost) AS Cost,SUM(A.AdoptNumPlan) AS AdoptNumPlan,SUM(A.AdoptNumResult) AdoptNumResult,CASE WHEN SUM(A.AdoptNumResult) > 0 THEN SUM(A.GroupCost) / SUM(A.AdoptNumResult) ELSE 0 END AS UnitCost,SUM(A.AdoptNumPlanPeriod) AS AdoptNumPlanPeriod FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,C.WorkingTypeCode,C.JobTypeCode,C.PrefectureCode,C.AdoptNum AS AdoptNumPlan,COALESCE(D.AdoptNum,0) AS AdoptNumResult,CASE WHEN COALESCE(B.AdoptNum,0) > 0 THEN A.Cost * CONVERT(FLOAT,C.AdoptNum) / CONVERT(FLOAT,B.AdoptNum) ELSE 0 END AS GroupCost,CONVERT(FLOAT,B.AdoptNum)*Period AS AdoptNumPlanPeriod FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.Cost * (CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1)) AS Cost,CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1) AS Period FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.EndDay,CONVERT(FLOAT,A.Cost) AS Cost,CASE WHEN A.StartDay >= @dStartDay THEN A.StartDay ELSE @dStartDay END AS StartDay2,CASE WHEN A.EndDay <= @dEndDay THEN A.EndDay ELSE @dEndDay END AS EndDay2 FROM CMPCostPerformanceMedia AS A WHERE (@vStartYM BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (@vEndYM BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (@vStartYM <= CONVERT(VARCHAR(6),A.StartDay,112) AND @vEndYM >= CONVERT(VARCHAR(6),A.EndDay,112))) AS A) AS A INNER JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptPlan AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay) AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay INNER JOIN CMPCostPerformanceAdoptPlan AS C ON A.CompanyCode = C.CompanyCode AND A.BranchSeq = C.BranchSeq AND A.MedName = C.MedName AND A.StartDay = C.StartDay LEFT JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE (A.AdoptMonth BETWEEN @vStartYM AND @vEndYM) AND A.AdoptMonth < CONVERT(VARCHAR(6),B.EndDay,112) GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode UNION ALL SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE A.AdoptMonth >= CONVERT(VARCHAR(6),B.EndDay,112) AND CONVERT(VARCHAR(6),B.EndDay,112) <= @vEndYM GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS D ON C.CompanyCode = D.CompanyCode AND C.BranchSeq = D.BranchSeq AND C.MedName = D.MedName AND C.StartDay = D.StartDay AND C.JobTypeCode = D.JobTypeCode AND C.WorkingTypeCode = D.WorkingTypeCode AND C.PrefectureCode = D.PrefectureCode) AS A "
		sSQL = sSQL & sWhere
		sSQL = sSQL & "GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay) AS A INNER JOIN CMPCostPerformanceBranch AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq ORDER BY SortNum,UnitCost;"

		sSQL2 = ""
		sSQL2 = sSQL2 & "/* �}�̏ڍׂ̕���E�X�ܕʎ��� */" & vbCrLf
		sSQL2 = sSQL2 & "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED" & vbCrLf
		sSQL2 = sSQL2 & "EXEC sp_executesql N'" & Replace(sSQL, "'", "''") & "'"
		If sDeclare <> "" Then sSQL2 = sSQL2 & vbCrLf & ",N'" & sDeclare & "'" & vbCrLf & sParams

		sqlMediaDetail = sSQL2 & vbCrLf
	End Function

	'******************************************************************************
	'�T�@�v�F�N�Ԑ��ڃf�[�^�擾SQL
	'���@���F
	'���@�l�F
	'���@���F2009/10/23 LIS K.Kokubo �쐬
	'******************************************************************************
	Function sqlYearChange()
		Dim sDeclare
		Dim sParams
		Dim sJoin
		Dim sWhere

		Dim sSQL
		Dim sSQL2
		Dim tmpYM1
		Dim tmpYM2

		sDeclare = ""
		sParams = ""
		sWhere = ""

		'�f�[�^�������`�F�b�N
		Call ChkData()

		tmpYM1 = YM1
		tmpYM2 = YM2
		If YM1 <> "" Then YM2 = Left(GetDateStr(DateAdd("d",-1,DateAdd("yyyy",1,CDate(Left(YM1,4) & "/" & Right(YM1,2) & "/01"))),""),6)
		Call setSQLData(sDeclare,sParams,sJoin,sWhere,"CompanyCode,BranchSeq,StartDay,EndDay,WorkingTypeCode,JobTypeCode,PrefectureCode,IndustryTypeCode,MedName")
		YM1 = tmpYM1
		YM2 = tmpYM2

		If sWhere <> "" Then
			sWhere = "WHERE " & sWhere
		End If

		sSQL = ""
		sSQL = sSQL & "DECLARE @idx TINYINT; DECLARE @tbl TABLE(YYYYMM CHAR(6) NOT NULL PRIMARY KEY,StartDay DATETIME NOT NULL,EndDay DATETIME NOT NULL); SET @idx = 0; WHILE @idx < 12 BEGIN INSERT INTO @tbl SELECT CONVERT(VARCHAR(6),DATEADD(MONTH,@idx,@dStartDay),112),DATEADD(MONTH,@idx,@dStartDay),DATEADD(DAY,-1,DATEADD(MONTH,@idx+1,@dStartDay)) SET @idx = @idx + 1; END; "
		sSQL = sSQL & "SELECT A.YYYYMM,A.CompanyCode,A.BranchSeq,A.BranchName,COALESCE(B.Cost,0) AS Cost,COALESCE(B.AdoptNumPlan,0) AS AdoptNumPlan,COALESCE(B.AdoptNumResult,0) AS AdoptNumResult,COALESCE(B.UnitCost,0) AS UnitCost FROM (SELECT DISTINCT A.YYYYMM,B.CompanyCode,B.BranchSeq,C.BranchName FROM @tbl AS A CROSS JOIN (SELECT A.CompanyCode,A.BranchSeq FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,B.WorkingTypeCode,B.JobTypeCode,B.PrefectureCode FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay FROM CMPCostPerformanceMedia AS A WHERE (@vStartYM BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (@vEndYM BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (@vStartYM <= CONVERT(VARCHAR(6),A.StartDay,112) AND @vEndYM >= CONVERT(VARCHAR(6),A.EndDay,112))) AS A) AS A INNER JOIN CMPCostPerformanceAdoptPlan AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,B.WorkingTypeCode,B.JobTypeCode,B.PrefectureCode) AS A "
		sSQL = sSQL & sWhere
		sSQL = sSQL & "GROUP BY A.CompanyCode,A.BranchSeq) AS B LEFT JOIN CMPCostPerformanceBranch AS C ON B.CompanyCode = C.CompanyCode AND B.BranchSeq = C.BranchSeq) AS A LEFT JOIN (SELECT A.CompanyCode,A.BranchSeq,A.YYYYMM,SUM(A.GroupCost) AS Cost,SUM(A.AdoptNumPlan) AS AdoptNumPlan,SUM(A.AdoptNumResult) AS AdoptNumResult,CASE WHEN SUM(A.AdoptNumResult) > 0 THEN SUM(A.GroupCost) / SUM(A.AdoptNumResult) ELSE 0 END AS UnitCost FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.YYYYMM,C.WorkingTypeCode,C.JobTypeCode,C.PrefectureCode,C.AdoptNum AS AdoptNumPlan,COALESCE(D.AdoptNum,0) AS AdoptNumResult,CASE WHEN COALESCE(B.AdoptNum,0) > 0 THEN A.Cost * CONVERT(FLOAT,C.AdoptNum) / CONVERT(FLOAT,B.AdoptNum) ELSE 0 END AS GroupCost FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.YYYYMM,(CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1)) AS Per,A.Cost * (CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1)) AS Cost FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.EndDay,CONVERT(FLOAT,A.Cost) AS Cost,CASE WHEN A.StartDay >= B.StartDay THEN A.StartDay ELSE B.StartDay END AS StartDay2,CASE WHEN A.EndDay <= B.EndDay THEN A.EndDay ELSE B.EndDay END AS EndDay2,B.YYYYMM FROM CMPCostPerformanceMedia AS A INNER JOIN @tbl AS B ON (B.StartDay BETWEEN A.StartDay AND A.EndDay) OR (B.EndDay BETWEEN A.StartDay AND A.EndDay) OR (B.StartDay <= A.StartDay AND B.EndDay >= A.EndDay) WHERE EXISTS(SELECT * FROM (SELECT @vCompanyCode AS CompanyCode UNION SELECT CompanyCode FROM CompanyInfo WHERE GroupCode = @vCompanyCode) AS TMP WHERE A.CompanyCode = TMP.CompanyCode)) AS A) AS A INNER JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptPlan AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay) AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay INNER JOIN CMPCostPerformanceAdoptPlan AS C ON A.CompanyCode = C.CompanyCode AND A.BranchSeq = C.BranchSeq AND A.MedName = C.MedName AND A.StartDay = C.StartDay LEFT JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,A.AdoptMonth,A.AdoptNum FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,A.AdoptMonth,A.AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE (A.AdoptMonth BETWEEN CONVERT(VARCHAR(6),@dStartDay,112) AND CONVERT(VARCHAR(6),@dEndDay,112)) AND A.AdoptMonth < CONVERT(VARCHAR(6),B.EndDay,112) UNION ALL SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,MIN(A.AdoptMonth) AS AdoptMonth,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE B.EndDay <= CONVERT(VARCHAR(6),@dEndDay,112) AND A.AdoptMonth >= CONVERT(VARCHAR(6),B.EndDay,112) GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS A) AS D ON C.CompanyCode = D.CompanyCode AND C.BranchSeq = D.BranchSeq AND C.MedName = D.MedName AND C.StartDay = D.StartDay AND C.JobTypeCode = D.JobTypeCode AND C.WorkingTypeCode = D.WorkingTypeCode AND C.PrefectureCode = D.PrefectureCode AND A.YYYYMM = D.AdoptMonth) AS A "
		sSQL = sSQL & sWhere
		sSQL = sSQL & "GROUP BY A.CompanyCode,A.BranchSeq,A.YYYYMM) AS B ON A.YYYYMM = B.YYYYMM AND A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq ORDER BY CompanyCode,BranchSeq,YYYYMM;"

		sSQL2 = ""
		sSQL2 = sSQL2 & "/*�i�r�E�̗p���P�T�|�[�g�V�X�e�� �N�Ԑ���*/" & vbCrLf
		sSQL2 = sSQL2 & "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;" & vbCrLf
		sSQL2 = sSQL2 & "SET NOCOUNT ON;" & vbCrLf
		sSQL2 = sSQL2 & "EXEC sp_executesql N'" & Replace(sSQL, "'", "''") & "'"
		If sDeclare <> "" Then sSQL2 = sSQL2 & vbCrLf & ",N'" & sDeclare & "'" & vbCrLf & sParams

		sqlYearChange = sSQL2 & vbCrLf
	End Function

	'******************************************************************************
	'�T�@�v�F�N�Ԑ��ڂ̓X�܁E����ꗗ
	'���@���F
	'���@�l�F
	'���@���F2010/02/10 LIS K.Kokubo �쐬
	'******************************************************************************
	Function sqlYearBranch()
		Dim sDeclare
		Dim sParams
		Dim sJoin
		Dim sWhere
		Dim tmpYM1
		Dim tmpYM2

		Dim sSQL
		Dim sSQL2

		sDeclare = ""
		sParams = ""
		sWhere = ""

		'�f�[�^�������`�F�b�N
		Call ChkData()

		tmpYM1 = YM1
		tmpYM2 = YM2
		If YM1 <> "" Then YM2 = Left(GetDateStr(DateAdd("d",-1,DateAdd("yyyy",1,CDate(Left(YM1,4) & "/" & Right(YM1,2) & "/01"))),""),6)
		Call setSQLData(sDeclare,sParams,sJoin,sWhere,"CompanyCode,BranchSeq,StartDay,EndDay,WorkingTypeCode,JobTypeCode,PrefectureCode,IndustryTypeCode,MedName")
		YM1 = tmpYM1
		YM2 = tmpYM2

		If sWhere <> "" Then
			sWhere = "WHERE " & sWhere
		End If

		sSQL = ""
		sSQL = sSQL & "SELECT A.CompanyCode,A.BranchSeq,B.BranchName FROM (SELECT A.CompanyCode,A.BranchSeq FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,B.WorkingTypeCode,B.JobTypeCode,B.PrefectureCode FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay FROM CMPCostPerformanceMedia AS A WHERE (@vStartYM BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (@vEndYM BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (@vStartYM <= CONVERT(VARCHAR(6),A.StartDay,112) AND @vEndYM >= CONVERT(VARCHAR(6),A.EndDay,112))) AS A) AS A INNER JOIN CMPCostPerformanceAdoptPlan AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,B.WorkingTypeCode,B.JobTypeCode,B.PrefectureCode) AS A "
		sSQL = sSQL & sWhere
		sSQL = sSQL & "GROUP BY A.CompanyCode,A.BranchSeq) AS A INNER JOIN CMPCostPerformanceBranch AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq;"

		sSQL2 = ""
		sSQL2 = sSQL2 & "/* �i�r�E�̗p���P�T�|�[�g�V�X�e�� �N�Ԑ��ڂ̓X�܁E����ꗗ */" & vbCrLf
		sSQL2 = sSQL2 & "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED" & vbCrLf
		sSQL2 = sSQL2 & "EXEC sp_executesql N'" & Replace(sSQL, "'", "''") & "'"
		If sDeclare <> "" Then sSQL2 = sSQL2 & vbCrLf & ",N'" & sDeclare & "'" & vbCrLf & sParams

		sqlYearBranch = sSQL2 & vbCrLf
	End Function

	'******************************************************************************
	'�T�@�v�F����E�X�ܖ��̔N�Ԑ��ڃf�[�^���擾����SQL
	'���@���F
	'���@�l�F
	'���@���F2009/10/27 LIS K.Kokubo �쐬
	'******************************************************************************
	Function sqlYearChangeBranch()
		Dim sDeclare
		Dim sParams
		Dim sJoin
		Dim sWhere

		Dim sSQL
		Dim sSQL2
		Dim tmp1

		sDeclare = ""
		sParams = ""
		sWhere = ""

		'�f�[�^�������`�F�b�N
		Call ChkData()

		tmp1 = YM1
		YM1 = ""
		YM2 = ""
		Call setSQLData(sDeclare,sParams,sJoin,sWhere,"CompanyCode.DspCompanyCode,BranchSeq,StartDay,EndDay,WorkingTypeCode,JobTypeCode,PrefectureCode,IndustryTypeCode,MedName")
		If tmp1 <> "" Then
			YM1 = tmp1

			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vStartYM VARCHAR(6)"
			sParams = sParams & ",@vStartYM = N'" & YM1 & "'"

			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@dStartDay DATETIME"
			sParams = sParams & ",@dStartDay = N'" & GetDateStr(CDate(Left(YM1,4) & "/" & Right(YM1,2) & "/01"),"") & "'"

			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@dEndDay DATETIME"
			sParams = sParams & ",@dEndDay = N'" & GetDateStr(DateAdd("d",-1,DateAdd("yyyy",1,CDate(Left(YM1,4) & "/" & Right(YM1,2) & "/01"))),"") & "'"
		End If

		If sWhere <> "" Then
			sWhere = "WHERE " & sWhere
		End If

		sSQL = ""
		sSQL = sSQL & "DECLARE @idx TINYINT; DECLARE @tbl TABLE(YYYYMM CHAR(6) NOT NULL PRIMARY KEY,StartDay DATETIME NOT NULL,EndDay DATETIME NOT NULL); SET @idx = 0; WHILE @idx < 12 BEGIN INSERT INTO @tbl SELECT CONVERT(VARCHAR(6),DATEADD(MONTH,@idx,@dStartDay),112),DATEADD(MONTH,@idx,@dStartDay),DATEADD(DAY,-1,DATEADD(MONTH,@idx+1,@dStartDay)) SET @idx = @idx + 1; END;"
		sSQL = sSQL & "SELECT A.YYYYMM,A.CompanyCode,A.BranchSeq,A.BranchName,A.MedName,COALESCE(B.Cost,0) AS Cost,COALESCE(B.AdoptNumPlan,0) AS AdoptNumPlan,COALESCE(B.AdoptNumResult,0) AS AdoptNumResult,COALESCE(B.UnitCost,0) AS UnitCost FROM (SELECT DISTINCT A.YYYYMM,B.CompanyCode,B.BranchSeq,B.MedName,C.BranchName FROM @tbl AS A CROSS JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName FROM CMPCostPerformanceMedia AS A WHERE EXISTS(SELECT * FROM (SELECT @vCompanyCode AS CompanyCode UNION SELECT CompanyCode FROM CompanyInfo WHERE GroupCode = @vCompanyCode) AS TMP WHERE A.CompanyCode = TMP.CompanyCode) AND A.CompanyCode = @vDspCompanyCode AND A.BranchSeq = @vBranchSeq AND ((@dStartDay BETWEEN A.StartDay AND A.EndDay) OR (@dEndDay BETWEEN A.StartDay AND A.EndDay) OR (@dStartDay <= A.StartDay AND @dEndDay >= A.EndDay))) AS B LEFT JOIN CMPCostPerformanceBranch AS C ON B.CompanyCode = C.CompanyCode AND B.BranchSeq = C.BranchSeq) AS A LEFT JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.YYYYMM,SUM(A.GroupCost) AS Cost,SUM(A.AdoptNumPlan) AS AdoptNumPlan,SUM(A.AdoptNumResult) AS AdoptNumResult,CASE WHEN SUM(A.AdoptNumResult) > 0 THEN SUM(A.GroupCost) / SUM(A.AdoptNumResult) ELSE 0 END AS UnitCost FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.YYYYMM,C.WorkingTypeCode,C.JobTypeCode,C.PrefectureCode,C.AdoptNum AS AdoptNumPlan,COALESCE(D.AdoptNum,0) AS AdoptNumResult,CASE WHEN COALESCE(B.AdoptNum,0) > 0 THEN A.Cost * CONVERT(FLOAT,C.AdoptNum) / CONVERT(FLOAT,B.AdoptNum) ELSE 0 END AS GroupCost FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.YYYYMM,(CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1)) AS Per,A.Cost * (CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1)) AS Cost FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.EndDay,CONVERT(FLOAT,A.Cost) AS Cost,CASE WHEN A.StartDay >= B.StartDay THEN A.StartDay ELSE B.StartDay END AS StartDay2,CASE WHEN A.EndDay <= B.EndDay THEN A.EndDay ELSE B.EndDay END AS EndDay2,B.YYYYMM FROM CMPCostPerformanceMedia AS A INNER JOIN @tbl AS B ON (B.StartDay BETWEEN A.StartDay AND A.EndDay) OR (B.EndDay BETWEEN A.StartDay AND A.EndDay) OR (B.StartDay <= A.StartDay AND B.EndDay >= A.EndDay) WHERE EXISTS(SELECT * FROM (SELECT @vCompanyCode AS CompanyCode UNION SELECT CompanyCode FROM CompanyInfo WHERE GroupCode = @vCompanyCode) AS TMP WHERE A.CompanyCode = TMP.CompanyCode)) AS A) AS A INNER JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptPlan AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay) AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay INNER JOIN CMPCostPerformanceAdoptPlan AS C ON A.CompanyCode = C.CompanyCode AND A.BranchSeq = C.BranchSeq AND A.MedName = C.MedName AND A.StartDay = C.StartDay LEFT JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,A.AdoptMonth,A.AdoptNum FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,A.AdoptMonth,A.AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE (A.AdoptMonth BETWEEN CONVERT(VARCHAR(6),@dStartDay,112) AND CONVERT(VARCHAR(6),@dEndDay,112)) AND A.AdoptMonth < CONVERT(VARCHAR(6),B.EndDay,112) UNION ALL SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,MIN(A.AdoptMonth) AS AdoptMonth,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE B.EndDay <= CONVERT(VARCHAR(6),@dEndDay,112) AND A.AdoptMonth >= CONVERT(VARCHAR(6),B.EndDay,112) GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS A) AS D ON C.CompanyCode = D.CompanyCode AND C.BranchSeq = D.BranchSeq AND C.MedName = D.MedName AND C.StartDay = D.StartDay AND C.JobTypeCode = D.JobTypeCode AND C.WorkingTypeCode = D.WorkingTypeCode AND C.PrefectureCode = D.PrefectureCode AND A.YYYYMM = D.AdoptMonth) AS A "
		sSQL = sSQL & sWhere
		sSQL = sSQL & "GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.YYYYMM) AS B ON A.YYYYMM = B.YYYYMM AND A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName ORDER BY CompanyCode,BranchSeq,MedName,YYYYMM;"

		sSQL2 = ""
		sSQL2 = sSQL2 & "/* �i�r�E�̗p���P�T�|�[�g�V�X�e�� �N�Ԑ��� */" & vbCrLf
		sSQL2 = sSQL2 & "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;" & vbCrLf
		sSQL2 = sSQL2 & "SET NOCOUNT ON;" & vbCrLf
		sSQL2 = sSQL2 & "EXEC sp_executesql N'" & Replace(sSQL, "'", "''") & "'"
		If sDeclare <> "" Then sSQL2 = sSQL2 & vbCrLf & ",N'" & sDeclare & "'" & vbCrLf & sParams

		sqlYearChangeBranch = sSQL2 & vbCrLf
	End Function

	'******************************************************************************
	'�T�@�v�F�V�~�����[�V�����f�t�H���g�f�[�^���oSQL�𐶐�
	'���@���F
	'���@�l�F
	'���@���F2009/10/23 LIS K.Kokubo �쐬
	'******************************************************************************
	Function sqlSimulation(ByVal vDefFlag)
		Dim sDeclare
		Dim sParams
		Dim sJoin
		Dim sWhere
		Dim tmpParam

		Dim sSQL
		Dim sSQL2

		sDeclare = ""
		sParams = ""
		sWhere = ""

		Y2 = Y1
		M2 = M1
		YM2 = YM1

		'�f�[�^�������`�F�b�N
		Call ChkData()

		Call setSQLData(sDeclare,sParams,sJoin,sWhere,"CompanyCode,BranchSeq,StartDay,EndDay,WorkingTypeCode,JobTypeCode,PrefectureCode,IndustryTypeCode,MedName")

		If sWhere <> "" Then
			sWhere = "WHERE " & sWhere
		End If

		tmpParam = ""
		If WorkingTypeCode <> "" Then: tmpParam = tmpParam & "@vWorkingTypeCode1": Else: tmpParam = tmpParam & "''": End If: tmpParam = tmpParam & ","
		If PrefectureCode <> "" Then: tmpParam = tmpParam & "@vPrefectureCode1": Else: tmpParam = tmpParam & "''": End If: tmpParam = tmpParam & ","
		If JobTypeCode <> "" Then: tmpParam = tmpParam & "@vJobTypeCode1": Else: tmpParam = tmpParam & "''": End If: tmpParam = tmpParam & ","
		If IndustryTypeCode <> "" Then: tmpParam = tmpParam & "@vIndustryTypeCode1": Else: tmpParam = tmpParam & "''": End If

		sSQL = ""
		sSQL = sSQL & "SELECT COALESCE(A.CompanyCode,B.CompanyCode) AS CompanyCode,COALESCE(A.BranchSeq,B.BranchSeq) AS BranchSeq,COALESCE(A.BranchName,B.BranchName) AS BranchName,A.Cost,A.AdoptNum,A.UnitCost,A.MedName,COALESCE(B.Cost,0) AS CostBef,COALESCE(B.AdoptNumPlan,0) AS AdoptPlanNumBef,COALESCE(B.AdoptNumResult,0) AS AdoptResultNumBef,COALESCE(B.UnitCost,0) AS UnitCostBef,COALESCE(B.AdoptNumPlanPeriod,0) AS AdoptPlanNumPeriodBef FROM ("
		If vDefFlag = "1" Then
			sSQL = sSQL & "SELECT A.CompanyCode,A.BranchSeq,A.BranchName,B.Cost,B.AdoptNumPlan AS AdoptNum,B.UnitCost,C.MedName,B.AdoptNumPlanPeriod FROM CMPCostPerformanceBranch AS A LEFT JOIN (SELECT A.CompanyCode,A.BranchSeq,SUM(A.GroupCost) AS Cost,SUM(A.AdoptNumPlan) AS AdoptNumPlan,SUM(A.AdoptNumResult) AS AdoptNumResult,CASE WHEN SUM(A.AdoptNumPlan) > 0 THEN SUM(A.GroupCost) / SUM(A.AdoptNumPlan) ELSE 0 END AS UnitCost,SUM(A.AdoptNumPlanPeriod) AS AdoptNumPlanPeriod FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,C.WorkingTypeCode,C.JobTypeCode,C.PrefectureCode,C.AdoptNum AS AdoptNumPlan,COALESCE(D.AdoptNum,0) AS AdoptNumResult,CASE WHEN COALESCE(B.AdoptNum,0) > 0 THEN A.Cost * CONVERT(FLOAT,C.AdoptNum) / CONVERT(FLOAT,B.AdoptNum) ELSE 0 END AS GroupCost,CONVERT(FLOAT,B.AdoptNum)*Period AS AdoptNumPlanPeriod FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.Cost * (CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1)) AS Cost,CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1) AS Period FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.EndDay,CONVERT(FLOAT,A.Cost) AS Cost,CASE WHEN A.StartDay >= @dStartDay THEN A.StartDay ELSE @dStartDay END AS StartDay2,CASE WHEN A.EndDay <= @dEndDay THEN A.EndDay ELSE @dEndDay END AS EndDay2 FROM CMPCostPerformanceMedia AS A WHERE (@vStartYM BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (@vEndYM BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (@vStartYM <= CONVERT(VARCHAR(6),A.StartDay,112) AND @vEndYM >= CONVERT(VARCHAR(6),A.EndDay,112))) AS A) AS A INNER JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptPlan AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay) AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay INNER JOIN CMPCostPerformanceAdoptPlan AS C ON A.CompanyCode = C.CompanyCode AND A.BranchSeq = C.BranchSeq AND A.MedName = C.MedName AND A.StartDay = C.StartDay LEFT JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE (A.AdoptMonth BETWEEN @vStartYM AND @vEndYM) AND A.AdoptMonth < CONVERT(VARCHAR(6),B.EndDay,112) GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode UNION ALL SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE A.AdoptMonth >= CONVERT(VARCHAR(6),B.EndDay,112) AND CONVERT(VARCHAR(6),B.EndDay,112) <= @vEndYM GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS D ON C.CompanyCode = D.CompanyCode AND C.BranchSeq = D.BranchSeq AND C.MedName = D.MedName AND C.StartDay = D.StartDay AND C.JobTypeCode = D.JobTypeCode AND C.WorkingTypeCode = D.WorkingTypeCode AND C.PrefectureCode = D.PrefectureCode) AS A "
			sSQL = sSQL & sWhere
			sSQL = sSQL & "GROUP BY A.CompanyCode,A.BranchSeq) AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq LEFT JOIN dbo.ftbl_CMPCostPerformanceBranch_Media(@vCompanyCode,@vStartYM,@vEndYM," & tmpParam & ",',') AS C ON A.CompanyCode = C.CompanyCode AND A.BranchSeq = C.BranchSeq WHERE A.CompanyCode = @vCompanyCode "
		Else
			sSQL = sSQL & "SELECT CPB.CompanyCode,CPB.BranchSeq,CPB.BranchName,COALESCE(CPS.Cost,0) AS Cost,COALESCE(CPS.AdoptNum,0) AS AdoptNum,CASE WHEN COALESCE(CPS.Cost,0) > 0 AND COALESCE(CPS.AdoptNum,0) > 0 THEN CONVERT(FLOAT,CPS.Cost) / CONVERT(FLOAT,CPS.AdoptNum) ELSE 0 END AS UnitCost,CPS.MedName FROM CMPCostPerformanceBranch AS CPB LEFT JOIN CMPCostPerformanceSimulation AS CPS ON CPB.CompanyCode = CPS.CompanyCode AND CPB.BranchSeq = CPS.BranchSeq AND CPS.SimulationMonth = @vStartYM WHERE CPB.CompanyCode = @vCompanyCode "
		End If
		sSQL = sSQL & ") AS A FULL JOIN (SELECT A.CompanyCode,A.BranchSeq,A.BranchName,B.Cost,B.AdoptNumPlan,B.AdoptNumResult,B.UnitCost,C.MedName,B.AdoptNumPlanPeriod FROM CMPCostPerformanceBranch AS A LEFT JOIN (SELECT A.CompanyCode,A.BranchSeq,SUM(A.GroupCost) AS Cost,SUM(A.AdoptNumPlan) AS AdoptNumPlan,SUM(A.AdoptNumResult) AS AdoptNumResult,CASE WHEN SUM(A.AdoptNumResult) > 0 THEN SUM(A.GroupCost) / SUM(A.AdoptNumResult) ELSE 0 END AS UnitCost,SUM(A.AdoptNumPlanPeriod) AS AdoptNumPlanPeriod FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,C.WorkingTypeCode,C.JobTypeCode,C.PrefectureCode,C.AdoptNum AS AdoptNumPlan,COALESCE(D.AdoptNum,0) AS AdoptNumResult,CASE WHEN COALESCE(B.AdoptNum,0) > 0 THEN A.Cost * CONVERT(FLOAT,C.AdoptNum) / CONVERT(FLOAT,B.AdoptNum) ELSE 0 END AS GroupCost,CONVERT(FLOAT,B.AdoptNum)*Period AS AdoptNumPlanPeriod FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.Cost * (CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1)) AS Cost,CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1) AS Period FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.EndDay,CONVERT(FLOAT,A.Cost) AS Cost,CASE WHEN A.StartDay >= DATEADD(MONTH,-1,@dStartDay) THEN A.StartDay ELSE DATEADD(MONTH,-1,@dStartDay) END AS StartDay2,CASE WHEN A.EndDay <= DATEADD(DAY,-1,@dStartDay) THEN A.EndDay ELSE DATEADD(DAY,-1,@dStartDay) END AS EndDay2 FROM CMPCostPerformanceMedia AS A WHERE (CONVERT(VARCHAR(6),DATEADD(MONTH,-1,@vStartYM+'01'),112) BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (CONVERT(VARCHAR(6),DATEADD(MONTH,-1,@vEndYM+'01'),112) BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (CONVERT(VARCHAR(6),DATEADD(MONTH,-1,@vStartYM+'01'),112) <= CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),DATEADD(MONTH,-1,@vEndYM+'01'),112) >= CONVERT(VARCHAR(6),A.EndDay,112))) AS A) AS A INNER JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptPlan AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay) AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay INNER JOIN CMPCostPerformanceAdoptPlan AS C ON A.CompanyCode = C.CompanyCode AND A.BranchSeq = C.BranchSeq AND A.MedName = C.MedName AND A.StartDay = C.StartDay LEFT JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE (A.AdoptMonth BETWEEN CONVERT(VARCHAR(6),DATEADD(MONTH,-1,@vStartYM+'01'),112) AND CONVERT(VARCHAR(6),DATEADD(MONTH,-1,@vEndYM+'01'),112)) AND A.AdoptMonth < CONVERT(VARCHAR(6),B.EndDay,112) GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode UNION ALL SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE A.AdoptMonth >= CONVERT(VARCHAR(6),B.EndDay,112) AND CONVERT(VARCHAR(6),B.EndDay,112) <= CONVERT(VARCHAR(8),DATEADD(MONTH,-1,@vEndYM+'01'),112) GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS D ON C.CompanyCode = D.CompanyCode AND C.BranchSeq = D.BranchSeq AND C.MedName = D.MedName AND C.StartDay = D.StartDay AND C.JobTypeCode = D.JobTypeCode AND C.WorkingTypeCode = D.WorkingTypeCode AND C.PrefectureCode = D.PrefectureCode) AS A WHERE A.CompanyCode = @vCompanyCode GROUP BY A.CompanyCode,A.BranchSeq) AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq LEFT JOIN dbo.ftbl_CMPCostPerformanceBranch_Media(@vCompanyCode,CONVERT(VARCHAR(8),DATEADD(MONTH,-1,@vStartYM+'01'),112),CONVERT(VARCHAR(8),DATEADD(MONTH,-1,@vEndYM+'01'),112)," & tmpParam & ",',') AS C ON A.CompanyCode = C.CompanyCode AND A.BranchSeq = C.BranchSeq WHERE A.CompanyCode = @vCompanyCode) AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq;"

		sSQL2 = ""
		sSQL2 = sSQL2 & "/* �i�r�E�̗p���P�T�|�[�g�V�X�e�� �V�~�����[�V���� */" & vbCrLf
		sSQL2 = sSQL2 & "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED" & vbCrLf
		sSQL2 = sSQL2 & "EXEC sp_executesql N'" & Replace(sSQL, "'", "''") & "'"
		If sDeclare <> "" Then sSQL2 = sSQL2 & vbCrLf & ",N'" & sDeclare & "'" & vbCrLf & sParams

		sqlSimulation = sSQL2 & vbCrLf
	End Function

	'******************************************************************************
	'�T�@�v�F�V�~�����[�V�����Q�l�f�[�^SQL�𐶐�
	'���@���F
	'���@�l�F
	'���@���F2009/11/20 LIS K.Kokubo �쐬
	'******************************************************************************
	Function sqlSimReferenceData()
		Dim sDeclare
		Dim sParams
		Dim sJoin
		Dim sWhere
		Dim sWhere2
		Dim tmp

		Dim sSQL
		Dim sSQL2
		Dim tmp1
		Dim iPrmNo

		Dim aValue
		Dim idx

		sDeclare = ""
		sParams = ""
		sWhere = ""

		If Y1 = "" Then
			If Month(Date) >= 4 And Month(Date) <= 12 Then
				Y1 = CStr(Year(Date))
			Else
				Y1 = CStr(Year(DateAdd("yyyy",-1,Date)))
			End If
		End If
		If M1 = "" Then M1 = CStr(Month(Date))
		YM1 = Y1 & Right("0"&M1,2)

		If Y2 = "" Or M2 = "" Or YM2 = "" Then
			Y2 = Y1
			M2 = M1
			YM2 = Left(GetDateStr(DateAdd("d",-1,DateAdd("m",1,CDate(Y1 & "/" & M1 & "/1"))),""),6)
		End If

		'�f�[�^�������`�F�b�N
		Call ChkData()

		Call setSQLData(sDeclare,sParams,tmp,tmp,"CompanyCode")
		sWhere = "NOT EXISTS(SELECT * FROM (SELECT @vCompanyCode AS CompanyCode UNION SELECT CompanyCode FROM CompanyInfo WHERE GroupCode = @vCompanyCode) AS TMP WHERE A.CompanyCode = TMP.CompanyCode) "

		Call setSQLData(sDeclare,sParams,sJoin,sWhere,"StartDay,EndDay,WorkingTypeCode,JobTypeCode,PrefectureCode,IndustryTypeCode,MedName")

		'<�Ζ��`��>
		tmp1 = ""
		iPrmNo = 1
		If WorkingTypeCode <> "" Then
			aValue = Split(WorkingTypeCode,",")
			For idx = 0 To UBound(aValue)
				If aValue(idx) <> "" Then
					If tmp1 <> "" Then tmp1 = tmp1 & ","
					tmp1 = tmp1 & "@vWorkingTypeCode" & iPrmNo

					iPrmNo = iPrmNo + 1
				End If
			Next

			If sWhere2 <> "" Then sWhere2 = sWhere2 & "AND "
			sWhere2 = sWhere2 & "A.WorkingTypeCode IN (" & tmp1 & ") "
		End If
		'<�Ζ��`��>

		'<�E��>
		tmp1 = ""
		iPrmNo = 1
		If JobTypeCode <> "" Then
			aValue = Split(JobTypeCode,",")
			For idx = 0 To UBound(aValue)
				If aValue(idx) <> "" Then
					If tmp1 <> "" Then tmp1 = tmp1 & ","
					tmp1 = tmp1 & "@vJobTypeCode" & iPrmNo

					iPrmNo = iPrmNo + 1
				End If
			Next

			If sWhere2 <> "" Then sWhere2 = sWhere2 & "AND "
			sWhere2 = sWhere2 & "A.JobTypeCode IN (" & tmp1 & ") "
		End If
		'</�E��>

		If sWhere <> "" Then
			sWhere = "WHERE " & sWhere
		End If

		If sWhere2 <> "" Then
			sWhere2 = "WHERE " & sWhere2
		End If

		sSQL = ""
		sSQL = sSQL & "SELECT 1 AS SortNum,CASE WHEN A.AdoptNumResult > 0 THEN 1 ELSE 2 END AS SortNum2,B.BranchName,A.MedName,A.Cost,A.AdoptNumPlan,A.AdoptNumResult,A.UnitCost,A.AdoptNumPlanPeriod FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,AVG(A.GroupCost) AS Cost,SUM(A.AdoptNumPlan) AS AdoptNumPlan,SUM(A.AdoptNumResult) AS AdoptNumResult,CASE WHEN SUM(A.AdoptNumResult) > 0 THEN SUM(A.GroupCost) / SUM(A.AdoptNumResult) ELSE 0 END AS UnitCost,SUM(A.AdoptNumPlanPeriod) AS AdoptNumPlanPeriod FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,C.WorkingTypeCode,C.JobTypeCode,C.PrefectureCode,C.AdoptNum AS AdoptNumPlan,COALESCE(D.AdoptNum,0) AS AdoptNumResult,CASE WHEN COALESCE(B.AdoptNum,0) > 0 THEN A.Cost * CONVERT(FLOAT,C.AdoptNum) / CONVERT(FLOAT,B.AdoptNum) ELSE 0 END AS GroupCost,CONVERT(FLOAT,B.AdoptNum)*Period AS AdoptNumPlanPeriod FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.Cost * (CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1)) AS Cost,CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1) AS Period FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.EndDay,CONVERT(FLOAT,A.Cost) AS Cost,CASE WHEN A.StartDay >= @dStartDay THEN A.StartDay ELSE @dStartDay END AS StartDay2,CASE WHEN A.EndDay <= @dEndDay THEN A.EndDay ELSE @dEndDay END AS EndDay2 FROM CMPCostPerformanceMedia AS A WHERE (@vStartYM BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (@vEndYM BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (@vStartYM <= CONVERT(VARCHAR(6),A.StartDay,112) AND @vEndYM >= CONVERT(VARCHAR(6),A.EndDay,112))) AS A) AS A INNER JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptPlan AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay) AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay INNER JOIN CMPCostPerformanceAdoptPlan AS C ON A.CompanyCode = C.CompanyCode AND A.BranchSeq = C.BranchSeq AND A.MedName = C.MedName AND A.StartDay = C.StartDay LEFT JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE (A.AdoptMonth BETWEEN @vStartYM AND @vEndYM) AND A.AdoptMonth < CONVERT(VARCHAR(6),B.EndDay,112) GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode UNION ALL SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE A.AdoptMonth >= CONVERT(VARCHAR(6),B.EndDay,112) AND CONVERT(VARCHAR(6),B.EndDay,112) <= @vEndYM GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS D ON C.CompanyCode = D.CompanyCode AND C.BranchSeq = D.BranchSeq AND C.MedName = D.MedName AND C.StartDay = D.StartDay AND C.JobTypeCode = D.JobTypeCode AND C.WorkingTypeCode = D.WorkingTypeCode AND C.PrefectureCode = D.PrefectureCode) AS A "
		sSQL = sSQL & sWhere2
		sSQL = sSQL & "GROUP BY A.CompanyCode,A.BranchSeq,A.MedName) AS A INNER JOIN CMPCostPerformanceBranch AS B ON EXISTS(SELECT * FROM (SELECT @vCompanyCode AS CompanyCode UNION SELECT CompanyCode FROM CompanyInfo WHERE GroupCode = @vCompanyCode) AS TMP WHERE B.CompanyCode = TMP.CompanyCode) AND A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq "
		sSQL = sSQL & "UNION ALL "
		sSQL = sSQL & "SELECT 2 AS SortNum,CASE WHEN SUM(A.AdoptNumResult) > 0 THEN 1 ELSE 2 END AS SortNum2,'����' AS BranchName,A.MedName,AVG(A.GroupCost) AS Cost,AVG(CONVERT(FLOAT,A.AdoptNumPlan)) AS AdoptNumPlan,AVG(CONVERT(FLOAT,A.AdoptNumResult)) AS AdoptNumResult,CASE WHEN SUM(A.AdoptNumResult) > 0 THEN SUM(A.GroupCost) / SUM(A.AdoptNumResult) ELSE 0 END AS UnitCost,SUM(A.AdoptNumPlanPeriod) AS AdoptNumPlanPeriod FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,C.WorkingTypeCode,C.JobTypeCode,C.PrefectureCode,C.AdoptNum AS AdoptNumPlan,COALESCE(D.AdoptNum,0) AS AdoptNumResult,CASE WHEN COALESCE(B.AdoptNum,0) > 0 THEN A.Cost * CONVERT(FLOAT,C.AdoptNum) / CONVERT(FLOAT,B.AdoptNum) ELSE 0 END AS GroupCost,CONVERT(FLOAT,B.AdoptNum)*Period AS AdoptNumPlanPeriod FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.Cost * (CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1)) AS Cost,CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1) AS Period FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.EndDay,CONVERT(FLOAT,A.Cost) AS Cost,CASE WHEN A.StartDay >= @dStartDay THEN A.StartDay ELSE @dStartDay END AS StartDay2,CASE WHEN A.EndDay <= @dEndDay THEN A.EndDay ELSE @dEndDay END AS EndDay2 FROM CMPCostPerformanceMedia AS A WHERE (@vStartYM BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (@vEndYM BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (@vStartYM <= CONVERT(VARCHAR(6),A.StartDay,112) AND @vEndYM >= CONVERT(VARCHAR(6),A.EndDay,112))) AS A) AS A INNER JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptPlan AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay) AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay INNER JOIN CMPCostPerformanceAdoptPlan AS C ON A.CompanyCode = C.CompanyCode AND A.BranchSeq = C.BranchSeq AND A.MedName = C.MedName AND A.StartDay = C.StartDay LEFT JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE (A.AdoptMonth BETWEEN @vStartYM AND @vEndYM) AND A.AdoptMonth < CONVERT(VARCHAR(6),B.EndDay,112) GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode UNION ALL SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE A.AdoptMonth >= CONVERT(VARCHAR(6),B.EndDay,112) AND CONVERT(VARCHAR(6),B.EndDay,112) <= @vEndYM GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS D ON C.CompanyCode = D.CompanyCode AND C.BranchSeq = D.BranchSeq AND C.MedName = D.MedName AND C.StartDay = D.StartDay AND C.JobTypeCode = D.JobTypeCode AND C.WorkingTypeCode = D.WorkingTypeCode AND C.PrefectureCode = D.PrefectureCode) AS A "
		sSQL = sSQL & sWhere
		sSQL = sSQL & "GROUP BY A.MedName ORDER BY SortNum ASC,SortNum2 ASC,UnitCost ASC;"

		sSQL2 = ""
		sSQL2 = sSQL2 & "/* �̗p���P�T�|�[�g�V�X�e�� �V�~�����[�V�����Q�l�f�[�^ */" & vbCrLf
		sSQL2 = sSQL2 & "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED" & vbCrLf
		sSQL2 = sSQL2 & "EXEC sp_executesql N'" & Replace(sSQL, "'", "''") & "'"
		If sDeclare <> "" Then sSQL2 = sSQL2 & vbCrLf & ",N'" & sDeclare & "'" & vbCrLf & sParams
		sqlSimReferenceData = sSQL2 & vbCrLf
	End Function


	'******************************************************************************
	'�T�@�v�F�����}�̔�rSQL�𐶐�
	'���@���F
	'���@�l�F
	'���@���F2009/11/06 LIS K.Kokubo �쐬
	'******************************************************************************
	Function sqlMedCompare()
		Dim sDeclare
		Dim sParams
		Dim sJoin
		Dim sWhere
		Dim sWhere2

		Dim sSQL
		Dim sSQL2
		Dim tmp1
		Dim iPrmNo

		Dim aValue
		Dim idx

		sDeclare = ""
		sParams = ""
		sWhere = ""

		If Y1 = "" Then
			If Month(Date) >= 4 And Month(Date) <= 12 Then
				Y1 = CStr(Year(Date))
			Else
				Y1 = CStr(Year(DateAdd("yyyy",-1,Date)))
			End If
		End If
		If M1 = "" Then M1 = CStr(Month(Date))
		YM1 = Y1 & Right("0"&M1,2)

		If Y2 = "" Or M2 = "" Or YM2 = "" Then
			Y2 = Y1
			M2 = M1
			YM2 = Left(GetDateStr(DateAdd("d",-1,DateAdd("m",1,CDate(Y1 & "/" & M1 & "/1"))),""),6)
		End If

		'�f�[�^�������`�F�b�N
		Call ChkData()

		Call setSQLData(sDeclare,sParams,sJoin,sWhere,"CompanyCode,BranchSeq,StartDay,EndDay,WorkingTypeCode,JobTypeCode,PrefectureCode,IndustryTypeCode,MedName")

		sWhere2 = sWhere2 & "NOT EXISTS(SELECT * FROM (SELECT @vCompanyCode AS CompanyCode UNION SELECT CompanyCode FROM CompanyInfo WHERE GroupCode = @vCompanyCode) AS TMP WHERE A.CompanyCode = TMP.CompanyCode) "

		'<�Ζ��`��>
		tmp1 = ""
		iPrmNo = 1
		If WorkingTypeCode <> "" Then
			aValue = Split(WorkingTypeCode,",")
			For idx = 0 To UBound(aValue)
				If aValue(idx) <> "" Then
					If tmp1 <> "" Then tmp1 = tmp1 & ","
					tmp1 = tmp1 & "@vWorkingTypeCode" & iPrmNo

					iPrmNo = iPrmNo + 1
				End If
			Next

			If sWhere2 <> "" Then sWhere2 = sWhere2 & "AND "
			sWhere2 = sWhere2 & "A.WorkingTypeCode IN (" & tmp1 & ") "
		End If
		'<�Ζ��`��>

		'<�E��>
		tmp1 = ""
		iPrmNo = 1
		If JobTypeCode <> "" Then
			aValue = Split(JobTypeCode,",")
			For idx = 0 To UBound(aValue)
				If aValue(idx) <> "" Then
					If tmp1 <> "" Then tmp1 = tmp1 & ","
					tmp1 = tmp1 & "@vJobTypeCode" & iPrmNo

					iPrmNo = iPrmNo + 1
				End If
			Next

			If sWhere2 <> "" Then sWhere2 = sWhere2 & "AND "
			sWhere2 = sWhere2 & "A.JobTypeCode IN (" & tmp1 & ") "
		End If
		'</�E��>

		'<�Ζ��n>
		tmp1 = ""
		iPrmNo = 1
		If PrefectureCode <> "" Then
			aValue = Split(PrefectureCode,",")
			For idx = 0 To UBound(aValue)
				If aValue(idx) <> "" Then
					If tmp1 <> "" Then tmp1 = tmp1 & ","
					tmp1 = tmp1 & "@vPrefectureCode" & iPrmNo

					iPrmNo = iPrmNo + 1
				End If
			Next

			If sWhere2 <> "" Then sWhere2 = sWhere2 & "AND "
			sWhere2 = sWhere2 & "A.PrefectureCode IN (" & tmp1 & ") "
		End If
		'</�Ζ��n>

		If sWhere <> "" Then
			sWhere = "WHERE " & sWhere
		End If

		If sWhere2 <> "" Then
			sWhere2 = "WHERE " & sWhere2
		End If

		sSQL = ""
		sSQL = sSQL & "SELECT A.MedName,A.AllCompanyCnt,A.SortNum,A.AllCostAvg,A.AllCost,A.AllAdoptNumPlan,A.AllAdoptNumResult,A.AllUnitCost,AllAdoptNumPlanPeriod,A.CompanyCnt,A.Cost,A.AdoptNumPlan,A.AdoptNumResult,A.UnitCost,AdoptNumPlanPeriod FROM (SELECT COALESCE(A.MedName,B.MedName) AS MedName,COALESCE(A.CompanyCnt,0) AS AllCompanyCnt,CASE WHEN B.MedName IS NOT NULL THEN 1 ELSE 2 END AS SortNum,COALESCE(A.AllCost,0)+COALESCE(B.AllCost,0) AS AllCost,COALESCE(A.AdoptNumPlan,0) AS AllAdoptNumPlan,COALESCE(A.AdoptNumResult,0) AS AllAdoptNumResult,COALESCE(A.UnitCost,0) AS AllUnitCost,COALESCE(B.Cost,0) AS Cost,COALESCE(B.AdoptNumPlan,0) AS AdoptNumPlan,COALESCE(B.AdoptNumResult,0) AS AdoptNumResult,COALESCE(B.UnitCost,0) AS UnitCost,(COALESCE(A.RangeCost,0)+COALESCE(B.RangeCost,0))/CONVERT(FLOAT,(COALESCE(A.CntMed,0)+COALESCE(B.CntMed,0)))/CONVERT(FLOAT,DATEDIFF(MONTH,@dStartDay,@dEndDay)+1) AS AllCostAvg,COALESCE(A.AdoptNumPlanPeriod,0)+COALESCE(B.AdoptNumPlanPeriod,0) AS AllAdoptNumPlanPeriod,COALESCE(B.AdoptNumPlanPeriod,0) AS AdoptNumPlanPeriod,COALESCE(B.CompanyCnt,0) AS CompanyCnt FROM (SELECT A.MedName,COUNT(DISTINCT A.CompanyCode) AS CompanyCnt,SUM(A.AdoptNumPlan) AS AdoptNumPlan,SUM(A.AdoptNumResult) AS AdoptNumResult,CASE WHEN SUM(A.AdoptNumResult) > 0 THEN SUM(A.GroupCost) / SUM(A.AdoptNumResult) ELSE 0 END AS UnitCost,SUM(A.Cost) AS AllCost,SUM(A.Days) AS AllDays,SUM(A.RangeCost) AS RangeCost,COUNT(*) AS CntMed,SUM(A.AdoptNumPlanPeriod) AS AdoptNumPlanPeriod FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.Cost,A.Days,C.WorkingTypeCode,C.JobTypeCode,C.PrefectureCode,C.AdoptNum AS AdoptNumPlan,COALESCE(D.AdoptNum,0) AS AdoptNumResult,A.Cost * CONVERT(FLOAT,C.AdoptNum) / CONVERT(FLOAT,B.AdoptNum) AS GroupCost,A.RangeCost,CONVERT(FLOAT,B.AdoptNum)*Period AS AdoptNumPlanPeriod FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1 AS Days,A.Cost * (CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1)) AS Cost,A.Cost / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1) * CONVERT(FLOAT,DATEDIFF(DAY,@dStartDay,@dEndDay)+1) AS RangeCost,CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1) AS Period FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.EndDay,CONVERT(FLOAT,A.Cost) AS Cost,CASE WHEN A.StartDay >= @dStartDay THEN A.StartDay ELSE @dStartDay END AS StartDay2,CASE WHEN A.EndDay <= @dEndDay THEN A.EndDay ELSE @dEndDay END AS EndDay2 FROM CMPCostPerformanceMedia AS A WHERE (@vStartYM BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (@vEndYM BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (@vStartYM <= CONVERT(VARCHAR(6),A.StartDay,112) AND @vEndYM >= CONVERT(VARCHAR(6),A.EndDay,112))) AS A) AS A INNER JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptPlan AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay) AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay INNER JOIN CMPCostPerformanceAdoptPlan AS C ON A.CompanyCode = C.CompanyCode AND A.BranchSeq = C.BranchSeq AND A.MedName = C.MedName AND A.StartDay = C.StartDay LEFT JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE (A.AdoptMonth BETWEEN @vStartYM AND @vEndYM) AND A.AdoptMonth < CONVERT(VARCHAR(6),B.EndDay,112) GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode UNION ALL SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE A.AdoptMonth >= CONVERT(VARCHAR(6),B.EndDay,112) AND CONVERT(VARCHAR(6),B.EndDay,112) <= @vEndYM GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS D ON C.CompanyCode = D.CompanyCode AND C.BranchSeq = D.BranchSeq AND C.MedName = D.MedName AND C.StartDay = D.StartDay AND C.JobTypeCode = D.JobTypeCode AND C.WorkingTypeCode = D.WorkingTypeCode AND C.PrefectureCode = D.PrefectureCode) AS A "
		sSQL = sSQL & sWhere2
		sSQL = sSQL & "GROUP BY A.MedName) AS A FULL JOIN (SELECT A.MedName,COUNT(DISTINCT A.CompanyCode) AS CompanyCnt,SUM(A.GroupCost) AS Cost,SUM(A.AdoptNumPlan) AS AdoptNumPlan,SUM(A.AdoptNumResult) AS AdoptNumResult,CASE WHEN SUM(A.AdoptNumResult) > 0 THEN SUM(A.GroupCost) / SUM(A.AdoptNumResult) ELSE 0 END AS UnitCost,SUM(A.Cost) AS AllCost,SUM(A.Days) AS AllDays,SUM(A.RangeCost) AS RangeCost,COUNT(*) AS CntMed,SUM(A.AdoptNumPlanPeriod) AS AdoptNumPlanPeriod FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.Cost,A.Days,C.WorkingTypeCode,C.JobTypeCode,C.PrefectureCode,C.AdoptNum AS AdoptNumPlan,COALESCE(D.AdoptNum,0) AS AdoptNumResult,A.Cost * CONVERT(FLOAT,C.AdoptNum) / CONVERT(FLOAT,B.AdoptNum) AS GroupCost,A.RangeCost,CONVERT(FLOAT,B.AdoptNum)*Period AS AdoptNumPlanPeriod FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1 AS Days,A.Cost * (CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1)) AS Cost,A.Cost / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1) * CONVERT(FLOAT,DATEDIFF(DAY,@dStartDay,@dEndDay)+1) AS RangeCost,CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1) AS Period FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.EndDay,CONVERT(FLOAT,A.Cost) AS Cost,CASE WHEN A.StartDay >= @dStartDay THEN A.StartDay ELSE @dStartDay END AS StartDay2,CASE WHEN A.EndDay <= @dEndDay THEN A.EndDay ELSE @dEndDay END AS EndDay2 FROM CMPCostPerformanceMedia AS A WHERE (@vStartYM BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (@vEndYM BETWEEN CONVERT(VARCHAR(6),A.StartDay,112) AND CONVERT(VARCHAR(6),A.EndDay,112)) OR (@vStartYM <= CONVERT(VARCHAR(6),A.StartDay,112) AND @vEndYM >= CONVERT(VARCHAR(6),A.EndDay,112))) AS A) AS A INNER JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptPlan AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay) AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay INNER JOIN CMPCostPerformanceAdoptPlan AS C ON A.CompanyCode = C.CompanyCode AND A.BranchSeq = C.BranchSeq AND A.MedName = C.MedName AND A.StartDay = C.StartDay LEFT JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE (A.AdoptMonth BETWEEN @vStartYM AND @vEndYM) AND A.AdoptMonth < CONVERT(VARCHAR(6),B.EndDay,112) GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode UNION ALL SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE A.AdoptMonth >= CONVERT(VARCHAR(6),B.EndDay,112) AND CONVERT(VARCHAR(6),B.EndDay,112) <= @vEndYM GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS D ON C.CompanyCode = D.CompanyCode AND C.BranchSeq = D.BranchSeq AND C.MedName = D.MedName AND C.StartDay = D.StartDay AND C.JobTypeCode = D.JobTypeCode AND C.WorkingTypeCode = D.WorkingTypeCode AND C.PrefectureCode = D.PrefectureCode) AS A "
		sSQL = sSQL & sWhere
		sSQL = sSQL & "GROUP BY A.MedName) AS B ON A.MedName = B.MedName) AS A ORDER BY A.SortNum,CASE WHEN A.SortNum = 1 AND COALESCE(A.UnitCost,0) = 0 THEN 2 WHEN A.SortNum = 2 AND COALESCE(A.AllUnitCost,0) = 0 THEN 2 ELSE 1 END ASC,CASE A.SortNum WHEN 1 THEN A.UnitCost WHEN 2 THEN A.AllUnitCost END ASC;"

		sSQL2 = ""
		sSQL2 = sSQL2 & "/* �̗p���P�T�|�[�g�V�X�e�� �����}�̔�r */" & vbCrLf
		sSQL2 = sSQL2 & "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED" & vbCrLf
		sSQL2 = sSQL2 & "EXEC sp_executesql N'" & Replace(sSQL, "'", "''") & "'"
		If sDeclare <> "" Then sSQL2 = sSQL2 & vbCrLf & ",N'" & sDeclare & "'" & vbCrLf & sParams
		sqlMedCompare = sSQL2 & vbCrLf
	End Function

	'******************************************************************************
	'�T�@�v�F�}�̔�r�N�Ԑ���SQL�𐶐�
	'���@���F
	'���@�l�F
	'���@���F2009/11/11 LIS K.Kokubo �쐬
	'******************************************************************************
	Function sqlMedCompareYear1()
		Dim sDeclare
		Dim sParams
		Dim sJoin
		Dim sWhere

		Dim sSQL
		Dim sSQL2
		Dim tmp1
		Dim tmpDate

		sDeclare = ""
		sParams = ""
		sWhere = ""

		'�f�[�^�������`�F�b�N
		Call ChkData()

		If Y1 = "" Then
			If Month(Date) >= 4 And Month(Date) <= 12 Then
				Y1 = CStr(Year(Date))
			Else
				Y1 = CStr(Year(DateAdd("yyyy",-1,Date)))
			End If
		End If
		If M1 = "" Then M1 = "4"
		YM1 = Y1 & Right("0"&M1,2)

		tmp1 = CompanyCode
		CompanyCode = ""
		If Y2 = "" Or M2 = "" Or YM2 = "" Then
			tmpDate = DateAdd("m",-1,DateAdd("yyyy",1,CDate(Y1 & "/" & Right("0"&M1,2) & "/01")))
			Y2 = Year(tmpDate)
			M2 = Month(tmpDate)
			YM2 = Y2 & Right("0"&M2,2)
		End If

		Call setSQLData(sDeclare,sParams,sJoin,sWhere,"CompanyCode,BranchSeq,StartDay,EndDay,WorkingTypeCode,JobTypeCode,PrefectureCode,IndustryTypeCode,MedName")

		CompanyCode = tmp1

		'<��ƃR�[�h>
		If CompanyCode <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vCompanyCode VARCHAR(8)"
			sParams = sParams & ",@vCompanyCode = N'" & CompanyCode & "'"
		End If
		'</��ƃR�[�h>

		If sWhere <> "" Then
			sWhere = "WHERE " & sWhere
		End If

		sSQL = ""
		sSQL = sSQL & "SELECT A.* INTO #TMP FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.YYYYMM,C.WorkingTypeCode,C.JobTypeCode,C.PrefectureCode,C.AdoptNum AS AdoptNumPlan,COALESCE(D.AdoptNum,0) AS AdoptNumResult,A.Cost * CONVERT(FLOAT,C.AdoptNum) / CONVERT(FLOAT,B.AdoptNum) AS GroupCost FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.YYYYMM,(CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1)) AS Per,A.Cost * (CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1)) AS Cost FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.EndDay,CONVERT(FLOAT,A.Cost) AS Cost,CASE WHEN A.StartDay >= B.StartDay THEN A.StartDay ELSE B.StartDay END AS StartDay2,CASE WHEN A.EndDay <= B.EndDay THEN A.EndDay ELSE B.EndDay END AS EndDay2,B.YYYYMM FROM CMPCostPerformanceMedia AS A INNER JOIN dbo.ftbl_Month(@vStartYM,@vEndYM) AS B ON (B.StartDay BETWEEN A.StartDay AND A.EndDay) OR (B.EndDay BETWEEN A.StartDay AND A.EndDay) OR (B.StartDay <= A.StartDay AND B.EndDay >= A.EndDay)) AS A) AS A INNER JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptPlan AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay) AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay INNER JOIN CMPCostPerformanceAdoptPlan AS C ON A.CompanyCode = C.CompanyCode AND A.BranchSeq = C.BranchSeq AND A.MedName = C.MedName AND A.StartDay = C.StartDay LEFT JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,A.AdoptMonth,A.AdoptNum FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,A.AdoptMonth,A.AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE (A.AdoptMonth BETWEEN CONVERT(VARCHAR(6),@vStartYM+'01',112) AND CONVERT(VARCHAR(6),DATEADD(DAY,-1,DATEADD(MONTH,1,@vEndYM+'01')),112)) AND A.AdoptMonth < CONVERT(VARCHAR(6),B.EndDay,112) UNION ALL SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,MIN(A.AdoptMonth) AS AdoptMonth,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE B.EndDay <= CONVERT(VARCHAR(6),DATEADD(DAY,-1,DATEADD(MONTH,1,@vEndYM+'01')),112) AND A.AdoptMonth >= CONVERT(VARCHAR(6),B.EndDay,112) GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS A) AS D ON C.CompanyCode = D.CompanyCode AND C.BranchSeq = D.BranchSeq AND C.MedName = D.MedName AND C.StartDay = D.StartDay AND C.JobTypeCode = D.JobTypeCode AND C.WorkingTypeCode = D.WorkingTypeCode AND C.PrefectureCode = D.PrefectureCode AND A.YYYYMM = D.AdoptMonth) AS A "
		sSQL = sSQL & sWhere
		sSQL = sSQL & ";" & vbCrLf
		sSQL = sSQL & "SELECT A.MedName,A.SortNum,[m1],[m2],[m3],[m4],[m5],[m6],[m7],[m8],[m9],[m10],[m11],[m12],COALESCE(B.UnitCostAvg,0) AS UnitCostAvg FROM (SELECT MedName,SortNum,[m1],[m2],[m3],[m4],[m5],[m6],[m7],[m8],[m9],[m10],[m11],[m12] FROM (SELECT '�������̑S�}�̕���' AS MedName,1 AS SortNum	,'m'+CONVERT(VARCHAR,DATEDIFF(MONTH,@vStartYM+'01',A.YYYYMM+'01')+1) AS MonthNum,CASE WHEN SUM(B.AdoptNumResult) > 0 THEN SUM(B.GroupCost) / SUM(B.AdoptNumResult) ELSE 0 END AS UnitCost FROM (SELECT B.MedName,A.YYYYMM,A.StartDay,A.EndDay FROM dbo.ftbl_Month(@vStartYM,@vEndYM) AS A CROSS JOIN (SELECT DISTINCT A.MedName FROM CMPCostPerformanceMedia AS A WHERE (@vStartYM+'01' BETWEEN A.StartDay AND A.EndDay) OR (DATEADD(DAY,-1,DATEADD(MONTH,1,@vEndYM+'01')) BETWEEN A.StartDay AND A.EndDay) OR (@vStartYM+'01' <= A.StartDay AND DATEADD(DAY,-1,DATEADD(MONTH,1,@vEndYM+'01')) >= A.EndDay)) AS B) AS A LEFT JOIN #TMP AS B ON A.YYYYMM = B.YYYYMM GROUP BY A.YYYYMM UNION ALL SELECT '���Ђ̕���' AS MedName,2 AS SortNum,'m'+CONVERT(VARCHAR,DATEDIFF(MONTH,@vStartYM+'01',A.YYYYMM+'01')+1) AS MonthNum,CASE WHEN SUM(B.AdoptNumResult) > 0 THEN SUM(B.GroupCost) / SUM(B.AdoptNumResult) ELSE 0 END AS UnitCost FROM (SELECT B.MedName,A.YYYYMM,A.StartDay,A.EndDay FROM dbo.ftbl_Month(@vStartYM,@vEndYM) AS A CROSS JOIN (SELECT DISTINCT A.MedName FROM CMPCostPerformanceMedia AS A WHERE (@vStartYM+'01' BETWEEN A.StartDay AND A.EndDay) OR (DATEADD(DAY,-1,DATEADD(MONTH,1,@vEndYM+'01')) BETWEEN A.StartDay AND A.EndDay) OR (@vStartYM+'01' <= A.StartDay AND DATEADD(DAY,-1,DATEADD(MONTH,1,@vEndYM+'01')) >= A.EndDay)) AS B) AS A LEFT JOIN (SELECT * FROM #TMP AS A WHERE EXISTS(SELECT * FROM (SELECT @vCompanyCode AS CompanyCode UNION SELECT CompanyCode FROM CompanyInfo WHERE GroupCode = @vCompanyCode) AS TMP WHERE A.CompanyCode = TMP.CompanyCode)) AS B ON A.YYYYMM = B.YYYYMM GROUP BY A.YYYYMM UNION ALL SELECT '���Ђ̂����ƃi�r' AS MedName,3 AS SortNum,'m'+CONVERT(VARCHAR,DATEDIFF(MONTH,@vStartYM+'01',A.YYYYMM+'01')+1) AS MonthNum,CASE WHEN SUM(B.AdoptNumResult) > 0 THEN SUM(B.GroupCost) / SUM(B.AdoptNumResult) ELSE 0 END AS UnitCost FROM (SELECT B.MedName,A.YYYYMM,A.StartDay,A.EndDay FROM dbo.ftbl_Month(@vStartYM,@vEndYM) AS A CROSS JOIN (SELECT DISTINCT A.MedName FROM CMPCostPerformanceMedia AS A WHERE (@vStartYM+'01' BETWEEN A.StartDay AND A.EndDay) OR (DATEADD(DAY,-1,DATEADD(MONTH,1,@vEndYM+'01')) BETWEEN A.StartDay AND A.EndDay) OR (@vStartYM+'01' <= A.StartDay AND DATEADD(DAY,-1,DATEADD(MONTH,1,@vEndYM+'01')) >= A.EndDay)) AS B) AS A LEFT JOIN (SELECT * FROM #TMP AS A WHERE EXISTS(SELECT * FROM (SELECT @vCompanyCode AS CompanyCode UNION SELECT CompanyCode FROM CompanyInfo WHERE GroupCode = @vCompanyCode) AS TMP WHERE A.CompanyCode = TMP.CompanyCode) AND A.MedName = '�����ƃi�r') AS B ON A.YYYYMM = B.YYYYMM GROUP BY A.YYYYMM/* UNION ALL SELECT '�����ƃi�r�i�S���ρj' AS MedName,4 AS SortNum,'m'+CONVERT(VARCHAR,DATEDIFF(MONTH,@vStartYM+'01',A.YYYYMM+'01')+1) AS MonthNum,CASE WHEN SUM(B.AdoptNumResult) > 0 THEN SUM(B.GroupCost) / SUM(B.AdoptNumResult) ELSE 0 END AS UnitCost FROM (SELECT B.MedName,A.YYYYMM,A.StartDay,A.EndDay FROM dbo.ftbl_Month(@vStartYM,@vEndYM) AS A CROSS JOIN (SELECT DISTINCT A.MedName FROM CMPCostPerformanceMedia AS A WHERE (@vStartYM+'01' BETWEEN A.StartDay AND A.EndDay) OR (DATEADD(DAY,-1,DATEADD(MONTH,1,@vEndYM+'01')) BETWEEN A.StartDay AND A.EndDay) OR (@vStartYM+'01' <= A.StartDay AND DATEADD(DAY,-1,DATEADD(MONTH,1,@vEndYM+'01')) >= A.EndDay)) AS B) AS A LEFT JOIN (SELECT * FROM #TMP AS A WHERE A.MedName = '�����ƃi�r') AS B ON A.YYYYMM = B.YYYYMM GROUP BY A.YYYYMM*/) AS A PIVOT (SUM(UnitCost) FOR A.MonthNum IN ([m1],[m2],[m3],[m4],[m5],[m6],[m7],[m8],[m9],[m10],[m11],[m12])) AS PVT) AS A LEFT JOIN (SELECT '�������̑S�}�̕���' AS MedName,CASE WHEN SUM(A.AdoptNumResult) > 0 THEN CONVERT(FLOAT,SUM(A.GroupCost)) / CONVERT(FLOAT,SUM(A.AdoptNumResult)) ELSE 0 END AS UnitCostAvg FROM #TMP AS A UNION ALL SELECT '���Ђ̕���' AS MedName,CASE WHEN SUM(A.AdoptNumResult) > 0 THEN CONVERT(FLOAT,SUM(A.GroupCost)) / CONVERT(FLOAT,SUM(A.AdoptNumResult)) ELSE 0 END AS UnitCostAvg FROM #TMP AS A WHERE EXISTS(SELECT * FROM (SELECT @vCompanyCode AS CompanyCode UNION SELECT CompanyCode FROM CompanyInfo WHERE GroupCode = @vCompanyCode) AS TMP WHERE A.CompanyCode = TMP.CompanyCode) UNION ALL SELECT '���Ђ̂����ƃi�r' AS MedName,CASE WHEN SUM(A.AdoptNumResult) > 0 THEN CONVERT(FLOAT,SUM(A.GroupCost)) / CONVERT(FLOAT,SUM(A.AdoptNumResult)) ELSE 0 END AS UnitCostAvg FROM #TMP AS A WHERE EXISTS(SELECT * FROM (SELECT @vCompanyCode AS CompanyCode UNION SELECT CompanyCode FROM CompanyInfo WHERE GroupCode = @vCompanyCode) AS TMP WHERE A.CompanyCode = TMP.CompanyCode) AND A.MedName = '�����ƃi�r') AS B ON A.MedName = B.MedName ORDER BY SortNum;"

		sSQL2 = ""
		sSQL2 = sSQL2 & "/* �̗p���P�T�|�[�g�V�X�e�� �}�̔�r�N�Ԑ��� */" & vbCrLf
		sSQL2 = sSQL2 & "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;SET NOCOUNT ON;" & vbCrLf
		sSQL2 = sSQL2 & "EXEC sp_executesql N'" & Replace(sSQL, "'", "''") & "'"
		If sDeclare <> "" Then sSQL2 = sSQL2 & vbCrLf & ",N'" & sDeclare & "'" & vbCrLf & sParams

		sqlMedCompareYear1 = sSQL2 & vbCrLf
	End Function

	'******************************************************************************
	'�T�@�v�F�X�܁E������}�̔�r�N�Ԑ���SQL�𐶐�
	'���@���F
	'���@�l�F
	'���@���F2009/11/10 LIS K.Kokubo �쐬
	'******************************************************************************
	Function sqlMedCompareYear2()
		Dim sDeclare
		Dim sParams
		Dim sJoin
		Dim sWhere

		Dim sSQL
		Dim sSQL2
		Dim tmp1
		Dim tmpDate

		sDeclare = ""
		sParams = ""
		sWhere = ""

		'�f�[�^�������`�F�b�N
		Call ChkData()

		If Y1 = "" Then
			If Month(Date) >= 4 And Month(Date) <= 12 Then
				Y1 = CStr(Year(Date))
			Else
				Y1 = CStr(Year(DateAdd("yyyy",-1,Date)))
			End If
		End If
		If M1 = "" Then M1 = "4"
		YM1 = Y1 & Right("0"&M1,2)

		tmp1 = CompanyCode
		CompanyCode = ""
		If Y2 = "" Or M2 = "" Or YM2 = "" Then
			tmpDate = DateAdd("m",-1,DateAdd("yyyy",1,CDate(Y1 & "/" & Right("0"&M1,2) & "/01")))
			Y2 = Year(tmpDate)
			M2 = Month(tmpDate)
			YM2 = Y2 & Right("0"&M2,2)
		End If

		Call setSQLData(sDeclare,sParams,sJoin,sWhere,"CompanyCode,BranchSeq,StartDay,EndDay,WorkingTypeCode,JobTypeCode,PrefectureCode,IndustryTypeCode,MedName")

		CompanyCode = tmp1

		If sWhere <> "" Then
			sWhere = "WHERE " & sWhere
		End If

		sSQL = ""
		sSQL = sSQL & "SELECT A.MedName,'m'+CONVERT(VARCHAR,DATEDIFF(MONTH,@vStartYM+'01',A.YYYYMM+'01')+1) AS MonthNum,COALESCE(B.UnitCost,0) AS UnitCost,COALESCE(B.Cost,0) AS Cost,COALESCE(B.AdoptNumPlan,0) AS AdoptNumPlan,COALESCE(B.AdoptNumResult,0) AS AdoptNumResult INTO #TMP FROM (SELECT B.MedName,A.YYYYMM,A.StartDay,A.EndDay FROM dbo.ftbl_Month(@vStartYM,@vEndYM) AS A CROSS JOIN (SELECT A.MedName FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,B.WorkingTypeCode,B.JobTypeCode,B.PrefectureCode FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay FROM CMPCostPerformanceMedia AS A WHERE (@vStartYM+'01' BETWEEN A.StartDay AND A.EndDay) OR (DATEADD(DAY,-1,DATEADD(MONTH,1,@vEndYM+'01')) BETWEEN A.StartDay AND A.EndDay) OR (@vStartYM+'01' <= A.StartDay AND DATEADD(DAY,-1,DATEADD(MONTH,1,@vEndYM+'01')) >= A.EndDay)) AS A) AS A INNER JOIN CMPCostPerformanceAdoptPlan AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,B.WorkingTypeCode,B.JobTypeCode,B.PrefectureCode) AS A "
		sSQL = sSQL & sWhere
		sSQL = sSQL & "GROUP BY A.MedName) AS B) AS A LEFT JOIN (SELECT A.YYYYMM,A.MedName,SUM(A.GroupCost) AS Cost,SUM(A.AdoptNumPlan) AS AdoptNumPlan,SUM(A.AdoptNumResult) AS AdoptNumResult,CASE WHEN SUM(A.AdoptNumResult) > 0 THEN SUM(A.GroupCost) / SUM(A.AdoptNumResult) ELSE 0 END AS UnitCost FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.YYYYMM,C.WorkingTypeCode,C.JobTypeCode,C.PrefectureCode,C.AdoptNum AS AdoptNumPlan,COALESCE(D.AdoptNum,0) AS AdoptNumResult,A.Cost * CONVERT(FLOAT,C.AdoptNum) / CONVERT(FLOAT,B.AdoptNum) AS GroupCost FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.YYYYMM,(CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1)) AS Per,A.Cost * (CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay2,A.EndDay2)+1) / CONVERT(FLOAT,DATEDIFF(DAY,A.StartDay,A.EndDay)+1)) AS Cost FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.EndDay,CONVERT(FLOAT,A.Cost) AS Cost,CASE WHEN A.StartDay >= B.StartDay THEN A.StartDay ELSE B.StartDay END AS StartDay2,CASE WHEN A.EndDay <= B.EndDay THEN A.EndDay ELSE B.EndDay END AS EndDay2,B.YYYYMM FROM CMPCostPerformanceMedia AS A INNER JOIN dbo.ftbl_Month(@vStartYM,@vEndYM) AS B ON (B.StartDay BETWEEN A.StartDay AND A.EndDay) OR (B.EndDay BETWEEN A.StartDay AND A.EndDay) OR (B.StartDay <= A.StartDay AND B.EndDay >= A.EndDay)) AS A) AS A INNER JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptPlan AS A GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay) AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay INNER JOIN CMPCostPerformanceAdoptPlan AS C ON A.CompanyCode = C.CompanyCode AND A.BranchSeq = C.BranchSeq AND A.MedName = C.MedName AND A.StartDay = C.StartDay LEFT JOIN (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,A.AdoptMonth,A.AdoptNum FROM (SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,A.AdoptMonth,A.AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE (A.AdoptMonth BETWEEN CONVERT(VARCHAR(6),@vStartYM+'01',112) AND CONVERT(VARCHAR(6),DATEADD(DAY,-1,DATEADD(MONTH,1,@vEndYM+'01')),112)) AND A.AdoptMonth < CONVERT(VARCHAR(6),B.EndDay,112) UNION ALL SELECT A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode,MIN(A.AdoptMonth) AS AdoptMonth,SUM(A.AdoptNum) AS AdoptNum FROM CMPCostPerformanceAdoptResult AS A INNER JOIN CMPCostPerformanceMedia AS B ON A.CompanyCode = B.CompanyCode AND A.BranchSeq = B.BranchSeq AND A.MedName = B.MedName AND A.StartDay = B.StartDay WHERE B.EndDay <= CONVERT(VARCHAR(6),DATEADD(DAY,-1,DATEADD(MONTH,1,@vEndYM+'01')),112) AND A.AdoptMonth >= CONVERT(VARCHAR(6),B.EndDay,112) GROUP BY A.CompanyCode,A.BranchSeq,A.MedName,A.StartDay,A.JobTypeCode,A.WorkingTypeCode,A.PrefectureCode) AS A) AS D ON C.CompanyCode = D.CompanyCode AND C.BranchSeq = D.BranchSeq AND C.MedName = D.MedName AND C.StartDay = D.StartDay AND C.JobTypeCode = D.JobTypeCode AND C.WorkingTypeCode = D.WorkingTypeCode AND C.PrefectureCode = D.PrefectureCode AND A.YYYYMM = D.AdoptMonth) AS A "
		sSQL = sSQL & sWhere
		sSQL = sSQL & "GROUP BY A.YYYYMM,A.MedName) AS B ON A.YYYYMM = B.YYYYMM AND A.MedName = B.MedName;" & vbCrLf

		sSQL = sSQL & "SELECT A.MedName,A.SortNum,[m1],[m2],[m3],[m4],[m5],[m6],[m7],[m8],[m9],[m10],[m11],[m12],CASE WHEN B.AdoptNumResult > 0 THEN B.Cost / B.AdoptNumResult ELSE 0 END AS UnitCostAvg FROM (SELECT MedName,CASE WHEN MedName = '�����ƃi�r' THEN 1 ELSE 2 END AS SortNum,[m1],[m2],[m3],[m4],[m5],[m6],[m7],[m8],[m9],[m10],[m11],[m12] FROM (SELECT A.MedName,A.MonthNum,A.UnitCost FROM #TMP AS A) AS A PIVOT (SUM(UnitCost) FOR A.MonthNum IN ([m1],[m2],[m3],[m4],[m5],[m6],[m7],[m8],[m9],[m10],[m11],[m12])) AS PVT) AS A LEFT JOIN (SELECT A.MedName,CONVERT(FLOAT,SUM(A.Cost)) AS Cost,CONVERT(FLOAT,SUM(A.AdoptNumResult)) AS AdoptNumResult FROM #TMP AS A GROUP BY A.MedName) AS B ON A.MedName = B.MedName ORDER BY SortNum,MedName;"

		sSQL2 = ""
		sSQL2 = sSQL2 & "/* �̗p���P�T�|�[�g�V�X�e�� �X�܁E������}�̔�r�N�Ԑ��� */" & vbCrLf
		sSQL2 = sSQL2 & "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED" & vbCrLf
		sSQL2 = sSQL2 & "EXEC sp_executesql N'" & Replace(sSQL, "'", "''") & "'"
		If sDeclare <> "" Then sSQL2 = sSQL2 & vbCrLf & ",N'" & sDeclare & "'" & vbCrLf & sParams

		sqlMedCompareYear2 = sSQL2 & vbCrLf
	End Function
End Class
%>
