<%
'**********************************************************************************************************************
'�T�@�v�F�l�ވꗗ /staff/person_result.asp
'�@�@�@�F�v���t�B�[�� /staff/person_detail.asp
'�@�@�@�F��L�y�[�W�ŏo�͗p�̊֐��Q�����̃t�@�C���ɗp�ӂ���B
'�@�@�@�F
'�@�@�@�F�������@�O������@������
'�@�@�@�F�v���O�C���N���[�h
'�@�@�@�F/config/personel.asp
'�@�@�@�F/include/commonfunc.asp
'��@���F�������@�o�͗p�@������
'�@�@�@�FGetHtmlNearbyStation	�F��{���y�[�W�̍Ŋ�w�ꗗ�g�s�l�k���擾
'�@�@�@�FGetHtmlResumeListBase	�F�w���E�E���E���i�ꗗ�̂P�s���̂g�s�l�k���擾
'�@�@�@�FGetHtmlSkillList		�F�X�L���y�[�W�ŃX�L���啪�ނP���̂g�s�l�k���擾
'**********************************************************************************************************************

'******************************************************************************
'�T�@�v�F�e�ҏW���ڂւ̃����N�{�^���g�s�l�k
'���@���F
'�g�p���F�����ƃi�r/staff/edit/edit1.asp
'�@�@�@�F�����ƃi�r/staff/edit/edit2.asp
'�@�@�@�F�����ƃi�r/staff/edit/edit3.asp
'�@�@�@�F�����ƃi�r/staff/edit/edit4.asp
'�@�@�@�F�����ƃi�r/staff/edit/edit5.asp
'�@�@�@�F�����ƃi�r/staff/edit/edit6.asp
'�@�@�@�F�����ƃi�r/staff/edit/edit7.asp
'�@�@�@�F�����ƃi�r/staff/edit/edit8.asp
'�@�@�@�F�����ƃi�r/staff/edit/edit9.asp
'���@�l�F
'�X�@�V�F2008/03/11 LIS K.Kokubo �쐬
'******************************************************************************
Function GetHtmlStaffEditList(ByVal vStaffCode, ByVal vCurrentURL)
	Dim sHTML
	Dim sDisabled

	sHTML = ""

	sDisabled = ""
	If InStr(vCurrentURL, "/staff/edit/edit1.asp") > 0 Then sDisabled = " disabled=""disabled"""
	sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit1.asp?staffcode=" & vStaffCode & """" & sDisabled & " style=""display:inline;""><input type=""submit"" value=""��{���""></form>"

	sDisabled = ""
	If InStr(vCurrentURL, "/staff/edit/edit2.asp") > 0 Then sDisabled = " disabled=""disabled"""
	sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit2.asp?staffcode=" & vStaffCode & """" & sDisabled & " style=""display:inline;""><input type=""submit"" value=""�w���E�E��""></form>"

	sDisabled = ""
	If InStr(vCurrentURL, "/staff/edit/edit3.asp") > 0 Then sDisabled = " disabled=""disabled"""
	sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit3.asp?staffcode=" & vStaffCode & """" & sDisabled & " style=""display:inline;""><input type=""submit"" value=""���i""></form>"

	sDisabled = ""
	If InStr(vCurrentURL, "/staff/edit/edit4.asp") > 0 Then sDisabled = " disabled=""disabled"""
	sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit4.asp?staffcode=" & vStaffCode & """" & sDisabled & " style=""display:inline;""><input type=""submit"" value=""�X�L��""></form>"

	sDisabled = ""
	If InStr(vCurrentURL, "/staff/edit/edit5.asp") > 0 Then sDisabled = " disabled=""disabled"""
	sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit5.asp?staffcode=" & vStaffCode & """" & sDisabled & " style=""display:inline;""><input type=""submit"" value=""IT�n�E��""></form>"

	sDisabled = ""
	If InStr(vCurrentURL, "/staff/edit/edit6.asp") > 0 Then sDisabled = " disabled=""disabled"""
	sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit6.asp?staffcode=" & vStaffCode & """" & sDisabled & " style=""display:inline;""><input type=""submit"" value=""��]����""></form>"

	sDisabled = ""
	If InStr(vCurrentURL, "/staff/edit/edit7.asp") > 0 Then sDisabled = " disabled=""disabled"""
	sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit7.asp?staffcode=" & vStaffCode & """" & sDisabled & " style=""display:inline;""><input type=""submit"" value=""�u�]���@""></form>"

	sDisabled = ""
	If InStr(vCurrentURL, "/staff/edit/edit8.asp") > 0 Then sDisabled = " disabled=""disabled"""
	sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit8.asp?staffcode=" & vStaffCode & """" & sDisabled & " style=""display:inline;""><input type=""submit"" value=""���ӕ��쓙""></form>"

	sDisabled = ""
	If InStr(vCurrentURL, "/staff/edit/edit9.asp") > 0 Then sDisabled = " disabled=""disabled"""
	sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit9.asp?staffcode=" & vStaffCode & """" & sDisabled & " style=""display:inline;""><input type=""submit"" value=""���Ȃo�q""></form>"

	sHTML = sHTML & "<br>"
	sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "staff/pdf/resume.asp?papersize=a3"" target=""_blank"">�i�h�r�������`�R</a>&nbsp;"
	sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "staff/pdf/resume.asp?papersize=a4"" target=""_blank"">�i�h�r�������`�S</a>&nbsp;"
	sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "staff/pdf/resume.asp?papersize=b4"" target=""_blank"">�������a�S</a>&nbsp;"
	sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "staff/pdf/resume.asp?papersize=b5"" target=""_blank"">�������a�T</a>&nbsp;"
	sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "staff/pdf/resume.asp?papersize=b42"" target=""_blank"">�o�C�g�������a�S</a>&nbsp;"
	sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "staff/pdf/experience.asp"" target=""_blank"">�E���o����</a>&nbsp;"
	sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "staff/pdf/experienceit.asp"" target=""_blank"">�h�s�n�E���o����</a>&nbsp;"

	GetHtmlStaffEditList = sHTML
End Function

'******************************************************************************
'�T�@�v�F��{���̍Ŋ�w�����̂g�s�l�k���擾
'���@���FrDB		�F�ڑ�����DB�R�l�N�V����
'�@�@�@�FvStaffCode	�F���E�҃R�[�h
'�g�p���F�����ƃi�r/staff/edit/edit1.asp
'���@�l�F
'�X�@�V�F2008/03/21 LIS K.Kokubo �쐬
'******************************************************************************
Function GetHtmlNearbyStation(ByRef rDB, ByVal vStaffCode)
	Dim sSQl
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbSeq
	Dim dbStationCode			'�Ŋ�w�F�w�R�[�h
	Dim dbStationName			'�Ŋ�w�F�w����
	Dim dbToStationBusFlag		'�Ŋ�w�܂ł̌�ʎ�i�F�o�X
	Dim dbToStationCarFlag		'�Ŋ�w�܂ł̌�ʎ�i�F��
	Dim dbToStationBicycleFlag	'�Ŋ�w�܂ł̌�ʎ�i�F���]��
	Dim dbToStationWalkFlag		'�Ŋ�w�܂ł̌�ʎ�i�F�k��
	Dim dbOtherTransportation	'�Ŋ�w�܂ł̌�ʎ�i�F���̑�
	Dim dbToStationTime			'�Ŋ�w�܂ł̎���

	Dim sHTML
	Dim sToStation
	Dim idx
	Dim flgDspAddButton

	sHTML = ""
	sToStation = ""
	flgDspAddButton = True

	sSQL = "sp_GetDataNearbyStation '" & G_USERID & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		If oRS.RecordCount > 1 Then flgDspAddButton = False
	End If

	'�u�ǉ��v�{�^��
	If flgDspAddButton = True Then
		sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit1_2.asp?flag=1"" method=""post"" style=""display:inline;"">"
		sHTML = sHTML & "<input type=""submit"" value=""�ǁ@��"">"
		sHTML = sHTML & "</form><br>" & vbCrLf
	Else
		sHTML = sHTML & "<input type=""button"" disabled=""disabled"" value=""�ǁ@��"" style=""margin:0px;"">&nbsp;"
		sHTML = sHTML & "<span style=""color:#333333;"">���Ŋ�w�͂Q�܂łł��B</span>" & vbCrLf
	End If

	idx = 0
	Do While GetRSState(oRS) = True And idx <= 1
		dbSeq = oRS.Collect("ID")
		dbStationCode = oRS.Collect("StationCode")
		dbStationName = oRS.Collect("StationName") & "�w"
		dbToStationBusFlag = oRS.Collect("ToStationBusFlag")
		dbToStationCarFlag = oRS.Collect("ToStationCarFlag")
		dbToStationBicycleFlag = oRS.Collect("ToStationBicycleFlag")
		dbToStationWalkFlag = oRS.Collect("ToStationWalkFlag")
		dbOtherTransportation = oRS.Collect("OtherTransportation")
		dbToStationTime = oRS.Collect("ToStationTime")

		sHTML = sHTML & "<div style=""padding:3px 0px; border-bottom:1px dotted #333333;"">"
		sHTML = sHTML & "<p class=""m0"" style=""float:left; width:350px;"">"
		sHTML = sHTML & dbStationName

		sToStation = ""
		If dbToStationBusFlag & dbToStationCarFlag & dbToStationBicycleFlag & dbToStationWalkFlag & dbOtherTransportation <> "" Then
			If dbToStationBusFlag = "1" Then
				If sToStation <> "" Then sToStation = sToStation & ",&nbsp;"
				sToStation = sToStation & "�o�X"
			End If

			If dbToStationCarFlag = "1" Then
				If sToStation <> "" Then sToStation = sToStation & ",&nbsp;"
				sToStation = sToStation & "��"
			End If

			If dbToStationBicycleFlag = "1" Then
				If sToStation <> "" Then sToStation = sToStation & ",&nbsp;"
				sToStation = sToStation & "���]��"
			End If

			If dbToStationWalkFlag = "1" Then
				If sToStation <> "" Then sToStation = sToStation & ",&nbsp;"
				sToStation = sToStation & "�k��"
			End If

			If dbOtherTransportation <> "" Then
				If sToStation <> "" Then sToStation = sToStation & ",&nbsp;"
				sToStation = sToStation & dbOtherTransportation
			End If
		End If

		If dbToStationTime <> "" And dbToStationTime > "0" Then
			sToStation = sToStation & "&nbsp;&nbsp;" & dbToStationTime & "��"
		End If

		If sToStation <> "" Then sToStation = "&nbsp;(" & sToStation & ")"

		sHTML = sHTML & sToStation
		sHTML = sHTML & "</p>" & vbCrLf

		sHTML = sHTML & "<div align=""right"" style=""float:right; width:140px;"">"
		sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit1_2.asp?flag=1&amp;seq=" & dbSeq & """ method=""post"" style=""display:inline;"">"
		sHTML = sHTML & "<input type=""submit"" value=""�ҁ@�W"">"
		sHTML = sHTML & "</form>&nbsp;"
		sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit1_2.asp?flag=0&amp;seq=" & dbSeq & """ method=""post"" style=""display:inline;"" onsubmit=""confirm('�u" & dbStationName & "�v���폜���܂����H');"">"
		sHTML = sHTML & "<input type=""submit"" value=""��@��"">"
		sHTML = sHTML & "</form><br>"
		sHTML = sHTML & "</div>" & vbCrLf

		sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
		sHTML = sHTML & "</div>" & vbCrLf

		oRS.MoveNext
		idx = idx + 1
	Loop
	Call RSClose(oRS)

	GetHtmlNearbyStation = sHTML
End Function

'******************************************************************************
'�T�@�v�F�w���E�E���ꗗ�̂P�s���̂g�s�l�k���擾
'���@���FvYear	�F�P��ځi�N�j
'�@�@�@�FvMonth	�F�Q��ځi���j
'�@�@�@�FvBody	�F�R��ځi�w����e�A�E����e�j
'�@�@�@�FvRow	�F�S��ځi�w��E���̍s���j
'�g�p���F�����ƃi�r/staff/edit/edit2.asp
'���@�l�F
'�X�@�V�F2008/02/28 LIS K.Kokubo �쐬
'******************************************************************************
Function GetHtmlResumeListBase(ByVal vYear, ByVal vMonth, ByVal vBody, ByVal vRow)
	Dim sHTML

	sHTML = ""
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td style=""vertical-align:top; font-size:10px;"">" & vRow & "�s</td>"
	sHTML = sHTML & "<td style=""border:1px solid #333333; text-align:right; vertical-align:top;"">" & vYear & "</td>"
	sHTML = sHTML & "<td style=""border:1px solid #333333; text-align:right; vertical-align:top;"">" & vMonth & "</td>"
	sHTML = sHTML & "<td style=""border:1px solid #333333; text-align:left; vertical-align:top;"">" & vBody & "</td>"
	sHTML = sHTML & "</tr>" & vbCrLf

	GetHtmlResumeListBase = sHTML
End Function

'******************************************************************************
'�T�@�v�F�X�L���啪�ނP���̂g�s�l�k���擾
'���@���FrDB			�F�ڑ����c�a�I�u�W�F�N�g
'�@�@�@�FvStaffCode		�F���O�C�������E�҃R�[�h
'�@�@�@�FvCategoryCode	�F�X�L���啪�ރR�[�h
'�g�p���F�����ƃi�r/staff/edit/edit4.asp
'���@�l�F
'�X�@�V�F2008/03/06 LIS K.Kokubo �쐬
'******************************************************************************
Function GetHtmlSkillList(ByRef rDB, ByVal vStaffCode, ByVal vCategoryCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbCategoryCode
	Dim dbSeq
	Dim dbCode
	Dim dbSkillName
	Dim dbStartDay
	Dim dbPeriod

	Dim sHTML
	Dim sCategoryName
	Dim iRecordCount
	Dim iMaxCount

	Select Case vCategoryCode
		Case "OS": sCategoryName = "�n�r": iMaxCount = 10
		Case "Application": sCategoryName = "�A�v���P�[�V����": iMaxCount = 10
		Case "DevelopmentLanguage": sCategoryName = "�J������": iMaxCount = 10
		Case "Database": sCategoryName = "�f�[�^�x�[�X": iMaxCount = 8
	End Select

	sHTML = ""

	sSQL = "sp_GetDataSkill '" & vStaffCode & "', '" & vCategoryCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

	iRecordCount = 0
	If GetRSState(oRS) = True Then
		'���R�[�h�Z�b�g�̐ؒf
		Set oRS.ActiveConnection = Nothing

		iRecordCount = oRS.RecordCount
	End If

	'�n�r�A�A�v���P�[�V�����A�J������͂P�O�܂œo�^�\
	'�f�[�^�x�[�X�͂W�܂œo�^�\
	If iRecordCount < iMaxCount Then
		sHTML = sHTML & "<form class=""m0"" action=""" & HTTPS_CURRENTURL & "staff/edit/edit4_1.asp?staffcode=" & vStaffCode & "&amp;flag=0&amp;categorycode=" & vCategoryCode & """ method=""post"" style=""display:inline;"">"
		sHTML = sHTML & "<input type=""submit"" value=""�ǁ@��"">"
		sHTML = sHTML & "</form>&nbsp;"
	Else
		sHTML = sHTML & "<input type=""submit"" disabled=""disabled"" value=""�ǁ@��"">&nbsp;"
	End If
	sHTML = sHTML & "��" & sCategoryName & "��" & iMaxCount & "�܂œo�^�ł��܂��B"

	If oRS.RecordCount = 0 Then
		sHTML = sHTML & "�o�^������܂���"
	End If

	Do While GetRSState(oRS) = True
		dbCategoryCode = oRS.Collect("CategoryCode")
		dbSeq = oRS.Collect("ID")
		dbCode = oRS.Collect("Code")
		dbSkillName = oRS.Collect("SkillName")
		dbStartDay = oRS.Collect("StartDay")
		dbPeriod = oRS.Collect("Period")

		sHTML = sHTML & "<div style=""margin-bottom:3px; border-bottom:1px dotted #999999;"">"
		sHTML = sHTML & "<div style=""float:left; width:350px;"">" & vbCrLf
		sHTML = sHTML & "<p class=""m0"">"
		sHTML = sHTML & "<span style=""background-color:#ccccff;"">"
		sHTML = sHTML & dbSkillName
		If dbPeriod <> "" Then sHTML = sHTML & "&nbsp;�g�p����(" & dbPeriod & "�N)"
		sHTML = sHTML & "</span>"
		sHTML = sHTML & "</p>" & vbCrLf
		sHTML = sHTML & "</div>" & vbCrLf
		sHTML = sHTML & "<div align=""right"" style=""float:left; width:140px;"">" & vbCrLf

		sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit4_1.asp?staffcode=" & vStaffCode & "&amp;flag=1&amp;categorycode=" & dbCategoryCode & "&amp;seq=" & dbSeq & """ method=""post"" style=""display:inline;"">"
		sHTML = sHTML & "<input type=""submit"" value=""�ҁ@�W"">"
		sHTML = sHTML & "</form>" & vbCrLf

		sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit4_1.asp?staffcode=" & vStaffCode & "&amp;flag=0&amp;categorycode=" & dbCategoryCode & "&amp;seq=" & dbSeq & """ method=""post"" style=""display:inline;"" onsubmit=""return confirm('�u" & dbSkillName & "�v���폜���܂����H');"">"
		sHTML = sHTML & "<input type=""submit"" value=""��@��"">"
		sHTML = sHTML & "</form>" & vbCrLf

		sHTML = sHTML & "</div>" & vbCrLf
		sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
		sHTML = sHTML & "</div>"

		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	GetHtmlSkillList = sHTML
End Function
%>
