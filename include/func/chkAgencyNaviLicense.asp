<%
'*******************************************************************************
'�T�@�v�F�㗝�X�戵�i�r���C�Z���X�`�F�b�N
'���@���F
'�߂�l�FBoolean
'���@�l�F
'�X�@�V�F2010/03/29 LIS K.Kokubo �쐬
'*******************************************************************************
Function chkAgencyNaviLicense(ByVal vApplicationCode,ByVal vAgencyCode,ByVal vBranchSeq)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sSQLErr

	Dim dbUserCode
	Dim dbHakouDate
	Dim dbRiyoFromDate
	Dim dbRiyoToDate
	Dim dbPlanTypeName
	Dim dbCompanyKbn
	Dim dbInterviewFlag
	Dim dbTempPermitFlag
	Dim dbIntroPermitFlag

	chkAgencyNaviLicense = False

	sSQL = "EXEC up_ChkAGC_MyNaviLicense '" & vApplicationCode & "','" & vAgencyCode & "','" & vBranchSeq & "';"
	flgQE = QUERYEXE(dbconn,oRS,sSQL,sSQLErr)
	If GetRSState(oRS) = True Then
		dbUserCode = oRS.Collect("UserCode")
		dbHakouDate = oRS.Collect("HakouDate")
		dbRiyoFromDate = oRS.Collect("RiyoFromDate")
		dbRiyoToDate = oRS.Collect("RiyoToDate")
		dbPlanTypeName = oRS.Collect("PlanTypeName")
		dbCompanyKbn = oRS.Collect("CompanyKbn")
		dbInterviewFlag = oRS.Collect("InterviewFlag")
		dbTempPermitFlag = oRS.Collect("TempPermitFlag")
		dbIntroPermitFlag = oRS.Collect("IntroPermitFlag")

		chkAgencyNaviLicense = True
	End If
	Call RSClose(oRS)

	G_USERID = dbUserCode
	G_USERTYPE = "company"
	G_APPLICATIONCODE = qsApplicationCode
	G_USEFLAG = "1"
	G_COMPANYKBN = dbCompanyKbn
	G_PLANTYPE = dbPlanTypeName
	G_INTERVIEWFLAG = dbInterviewFlag
	G_TEMPPERMITFLAG = dbTempPermitFlag
	G_INTROPERMITFLAG = dbIntroPermitFlag
	If dbRiyoFromDate <= Date And dbRiyoToDate >= Date Then
		G_PUBLICFLAG = "1"
	Else
		G_PUBLICFLAG = "0"
	End If
End Function
%>
