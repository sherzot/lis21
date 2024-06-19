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
'��@���F�������@���ʁ@������
'�@�@�@�FGetStrLastAccessDay				�F�ŏI�A�N�Z�X���̕������擾
'�@�@�@�F�������@�l�ވꗗ�y�[�W�o�͗p�@������
'�@�@�@�FGetHtmlPageControl					�F���E�҈ꗗ�y�[�W�̃y�[�W�R���g���[���g�s�l�k���擾
'�@�@�@�FDspStaffOne						�F���E�҈ꗗ�\���̂P�l���̘g�𐶐�
'�@�@�@�F�������@�v���t�B�[���y�[�W�o�͗p�@������
'�@�@�@�FDspProfileBase						�F�v���t�B�[���y�[�W�̊�{��񕔕����o��
'�@�@�@�FDspProfileNearbyStation			�F�v���t�B�[���y�[�W�̍Ŋ�w�������o��
'�@�@�@�FDspProfileEducateHistory			�F�v���t�B�[���y�[�W�̊w����񕔕����o��
'�@�@�@�FDspProfileCareerHistory			�F�v���t�B�[���y�[�W�̐E����񕔕����o��
'�@�@�@�FDspProfileCareerHistoryIT			�F�v���t�B�[���y�[�W�̂h�s�E����񕔕����o��
'�@�@�@�FDspProfileSkill					�F�v���t�B�[���y�[�W�̃X�L���������o��
'�@�@�@�FDspProfileSkillSimple					�F�v���t�B�[���y�[�W�̃X�L���������o��(�ȈՓo�^)
'�@�@�@�FDspProfileHope						�F�v���t�B�[���y�[�W�̊�]�����������o��
'�@�@�@�FDspProfileHopeWorkingType			�F�v���t�B�[���y�[�W�̊�]�����Ζ��`�ԕ������o��
'�@�@�@�FDspProfileHopeIndustry				�F�v���t�B�[���y�[�W�̊�]�����Ǝ핔�����o��
'�@�@�@�FDspProfileHopeJobType				�F�v���t�B�[���y�[�W�̊�]�����Ǝ핔�����o��
'�@�@�@�FDspProfileHopeWorkingPlace			�F�v���t�B�[���y�[�W�̊�]�����Ζ��n�������o��
'�@�@�@�FDspProfileHopeSalary				�F�v���t�B�[���y�[�W�̊�]�������^�������o��
'�@�@�@�FDspProfileHopeSpan					�F�v���t�B�[���y�[�W�̊�]�������ԁE���ԕ������o��
'�@�@�@�FDspProfileHopeWelfare				�F�v���t�B�[���y�[�W�̊�]�������������������o��
'�@�@�@�FDspCareerAnalyzer					�F�v���t�B�[���y�[�W�̏����I�ȗ��z�����o��
'�@�@�@�FDspProfileMail						�F�v���t�B�[���y�[�W�̍ŐV���M���[���󋵕������o��
'�@�@�@�FDspProfileStaffCode				�F�v���t�B�[���y�[�W�̋��E�҃R�[�h���o��
'�@�@�@�FDspProfileEditButton				�F�v���t�B�[���y�[�W�̊e�ҏW�{�^�����o��
'�@�@�@�FDspProfileUpdateDay				�F�v���t�B�[���y�[�W�̍ŏI�X�V�����o��
'�@�@�@�FDspProfileAttentionCareerHistory	�F�v���t�B�[���y�[�W�̐E����͖����ɑ΂��镶�����o��
'�@�@�@�F�������@�c�a�������݁@������
'�@�@�@�FRegMailMagazineAccess				�F���[���}�K�W������̃v���t�B�[���y�[�W�ւ̃A�N�Z�X�����O�ɋL�^
'�@�@�@�FUpdAccessCount						�F�v���t�B�[���y�[�W�̃A�N�Z�X�񐔂̃J�E���g�A�b�v
'�@�@�@�FSendMailStaffEdit					�F���X�V�̒ʒm�����X�̎Ј��Ƀ��[������
'�@�@�@�FRegAccessHistoryStaff				�F��Ƃ����E�ҏڍׂ��{�������烍�O�ɏ�������
'�@�@�@�FRegAccessHistoryStaffList			�F��Ƃ����E�҈ꗗ���{�������烍�O�ɏ�������
'**********************************************************************************************************************

'******************************************************************************
'�T�@�v�F�ŏI�A�N�Z�X���̕������擾
'���@���FvLastAccessDay	�F�ŏI�A�N�Z�X��
'���@�l�F
'���@���F2009/04/27 LIS K.Kokubo �쐬
'�@�@�@�F2009/06/23 LIS K.Kokubo �u�R�`�U�����A�N�Z�X�Ȃ��v��ǉ�
'******************************************************************************
Function GetStrLastAccessDay(ByVal vLastAccessDay)
	Dim sTxt

	If DateAdd("m", -3, Date) <= vLastAccessDay Then
		sTxt = GetDateStr(vLastAccessDay, "/")
	ElseIf DateAdd("m", -6, Date) <= vLastAccessDay Then
		sTxt = GetDateStr(vLastAccessDay, "/") & "�i�R�`�U�����A�N�Z�X�Ȃ��j"
	Else
		sTxt = "�U�����ȏ�A�N�Z�X�Ȃ�"
	End If
	'<���O�C�����܂�ߕ\��Ver.>
'	If DateDiff("d", rRS.Collect("LastAccessDay"), Date) <= 2 Then
'		sLastAccess = "�R���ȓ�"
'	ElseIf DateDiff("d", rRS.Collect("LastAccessDay"), Date) <= 6 Then
'		sLastAccess = "�V���ȓ�"
'	ElseIf DateDiff("m", rRS.Collect("LastAccessDay"), Date) = 0 Then
'		sLastAccess = "�P�����ȓ�"
'	ElseIf DateDiff("m", rRS.Collect("LastAccessDay"), Date) <= 2 Then
'		sLastAccess = "�R�����ȓ�"
'	ElseIf DateDiff("m", rRS.Collect("LastAccessDay"), Date) <= 5 Then
'		sLastAccess = "���N�ȓ�"
'	ElseIf DateDiff("m", rRS.Collect("LastAccessDay"), Date) <= 11 Then
'		sLastAccess = "�P�N�ȓ�"
'	Else
'		sLastAccess = "�P�N�ȏ�~"
'	End If
	'</���O�C�����܂�ߕ\��Ver.>

	GetStrLastAccessDay = sTxt
End Function

'******************************************************************************
'�T�@�v�F���E�҈ꗗ�y�[�W�̃y�[�W�R���g���[���g�s�l�k���擾
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�F���E�Ҍ������ʂ�ێ�����̃��R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvPageSize		�F�P�y�[�W������̕\������
'�@�@�@�FvPage			�F�\�����y�[�W
'�g�p���F�����ƃi�r/staff/person_list.asp
'�@�@�@�F�����ƃi�r/order/company_order.asp
'���@�l�F
'�X�@�V�F2007/02/11 LIS K.Kokubo �쐬
'******************************************************************************
Function GetHtmlPageControl(ByRef rDB, ByRef rRS, ByVal vPageSize, ByVal vPage)
	Dim iMaxPage
	Dim iLine
	Dim S_Page
	Dim E_Page
	Dim Sort
	Dim idx

	If GetRSState(rRS) = False Then Exit Function

	If vPage <> "" Then vPage = CInt(vPage)

	'�y�[�W������̕\������
	rRS.PageSize = vPageSize

	iMaxPage = rRS.PageCount
	If vPage > iMaxPage Then vPage = iMaxPage
	rRS.AbsolutePage = vPage

	'��ʏ�ɕ\������J�n�E�I���y�[�W�ԍ���ݒ�
	'�\���J�n�y�[�W�ԍ����w��
	S_Page = vPage - 5
	If S_Page < 1 Then
		S_Page = 1
	End If

	'�\���I���y�[�W�ԍ����w��
	E_Page = vPage + 4
	If E_Page < 10 Then E_Page = 10
	If E_Page > iMaxPage Then
		E_Page = iMaxPage
	End If
	If S_Page > iMaxPage - 9 And iMaxPage - 9 > 0 Then S_Page = iMaxPage - 9

	GetHtmlPageControl = ""
	GetHtmlPageControl = GetHtmlPageControl & "<table class=""cw"" style=""margin:10px 0px;"">"
	GetHtmlPageControl = GetHtmlPageControl & "<tbody>"
	GetHtmlPageControl = GetHtmlPageControl & "<tr>"
	GetHtmlPageControl = GetHtmlPageControl & "<td style=""width:88px; padding:5px; border:1px dotted #666699; border-width:1px 0px 1px 1px; text-align:center; background-color:#e8e8ff;"">"

	If vPage > 1 Then GetHtmlPageControl = GetHtmlPageControl & "<a href='javascript:ChgPage(" & vPage - 1 & ");'>�O�̃y�[�W</a>"
	GetHtmlPageControl = GetHtmlPageControl & "</td>"
	GetHtmlPageControl = GetHtmlPageControl & "<td style=""width:389px; padding:5px; border:1px dotted #666699; border-width:1px 0px 1px 0px; text-align:center; background-color:#e8e8ff;"">"

	If S_Page <> 1 Then GetHtmlPageControl = GetHtmlPageControl & "�c"

	For idx = S_Page To E_Page	'�y�[�W�ԍ���\��
		GetHtmlPageControl = GetHtmlPageControl & "�@"
		If idx = vPage Then		'�w��y�[�W�̕\��
			GetHtmlPageControl = GetHtmlPageControl & "[" & idx & "]"
		Else
			GetHtmlPageControl = GetHtmlPageControl & "<a href='javascript:ChgPage(" & idx & ");'>" & idx & "</a>"
		End If
	Next

	If E_Page < iMaxPage Then GetHtmlPageControl = GetHtmlPageControl & "�@�c"

	GetHtmlPageControl = GetHtmlPageControl & "</td>"
	GetHtmlPageControl = GetHtmlPageControl & "<td style=""width:89px; padding:5px; border:1px dotted #666699; border-width:1px 1px 1px 0px; text-align:center; background-color:#e8e8ff;"">"

	If vPage < iMaxPage Then GetHtmlPageControl = GetHtmlPageControl & "<a href='javascript:ChgPage(" & vPage + 1 & ");'>���̃y�[�W</a>"

	GetHtmlPageControl = GetHtmlPageControl & "</td>"
	GetHtmlPageControl = GetHtmlPageControl & "</tr>"
	GetHtmlPageControl = GetHtmlPageControl & "</tbody>"
	GetHtmlPageControl = GetHtmlPageControl & "</table>"
End Function

'******************************************************************************
'�T�@�v�F���E�҈ꗗ�\���̂P�l���̘g�𐶐�
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�F���E�Ҍ������ʂ�ێ�����̃��R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [vUserID]
'�g�p���F�i�r/staff/person_list.asp
'���@�l�F
'���@���F2007/04/09 LIS K.Kokubo �쐬
'�@�@�@�F2009/04/27 LIS K.Kokubo �ŏI�A�N�Z�X����\��
'�@�@�@�F2015/10/29 LIS K.Kimura �l���ی�̂��߁A�X�^�b�t�R�[�h�A�ŏI�A�N�Z�X���A����PR���\���ɕύX
'******************************************************************************
Function DspStaffOne(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbViewStaffDay		'���E�҉{����

	Dim sStaffCode
	Dim sClassName
	Dim sNewHtml			'�V���A�C�R���g�s�l�k
	Dim sHotHtml			'�g�n�s�A�C�R���g�s�l�k
	Dim sWorkerAlarmHtml	'WorkerAlarm�A�C�R���g�s�l�k
	Dim sScoutHtml			'���[�����M�ς݁i���Č��܂ށj�A�C�R���g�s�l�k
	Dim sViewStaffHtml		'�{���ς݁i���Č��܂ށj�A�C�R���g�s�l�k
	Dim sLastAccess			'���O�C�����\��
	Dim wSelfPR
	Dim wRow
	Dim sStaffDetailURL

	'�ϐ������� start
	sNewHtml = ""
	sHotHtml = ""
	sWorkerAlarmHtml = ""
	sScoutHtml = ""
	sViewStaffHtml = ""
	'�ϐ������� end

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")
	If vUserType = "staff" Or vUserType = "" Then
		sClassName = "pattern9"
	Else
		sClassName = "pattern8"
	End If

	If InStr(LCase(Request.ServerVariables("URL")), "jinzai") <> 0 Then
		'�l�ލ̗p�y�[�W�̋��E�ҏڍׂt�q�k
		sStaffDetailURL = "./person_detail.asp?staffcode=" & sStaffCode
	Else
		'�����ƃi�r�̋��E�ҏڍׂt�q�k
		sStaffDetailURL = "/staff/person_detail.asp?staffcode=" & sStaffCode & "&amp;ordercode=" & vOrderCode
	End If

	If sStaffCode <> "" Then
		wRow = 4
		'����PR
		wSelfPR = ""
		sSQL = "EXEC up_DtlStaffList '" & sStaffCode & "', '" & vUserID & "', '" & vOrderCode & "';"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			Set oRS.ActiveConnection = Nothing

			dbViewStaffDay = oRS.Collect("ViewStaffDay")

			'���Ȃq�q�s�ǉ�
			'If ChkStr(oRS.Collect("SelfPR")) <> "" Then
			'	wSelfPR = oRS.Collect("SelfPR")
			'	wRow = wRow + 1
			'End If

			'�E���s�ǉ�
			If ChkStr(oRS.Collect("CareerJobType")) <> "" Then
				wRow = wRow + 1
			End If

			'<�A�C�R���ݒ�>
			'�V��
			If oRS.Collect("NewFlag") = "1" Then sNewHtml = "<a href=""javascript:void(0)"" onclick='window.open(""/staff/icon_navi.html"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=420,height=300"")'><img src=""/img/c_new.gif"" alt="""" border=""0""></a>&nbsp;"
			'�{���񐔑���
			If oRS.Collect("HotFlag") = "1" Then sHotHtml = "<a href=""javascript:void(0)"" onclick='window.open(""/staff/icon_navi.html"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=420,height=300"")'><img src=""/img/c_hot.gif"" alt="""" border=""0""></a>&nbsp;"
			'WorkerAlarm
			If oRS.Collect("WorkerAlarmFlag") = "1" Then sWorkerAlarmHtml = "<a href=""javascript:void(0)"" onclick='window.open(""/staff/icon_navi.html"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=420,height=300"")'><img src=""/img/workeralarm.gif"" alt=""�����ɓ����鋁�E�҃}�[�N"" border=""0""></a>&nbsp;"
			If oRS.Collect("ViewStaffFlag") = "1" Then
				'�{���ς�
				sViewStaffHtml = "<a href=""javascript:void(0)"" onclick='window.open(""/staff/icon_navi.html"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=420,height=300"")'><img src=""/img/viewstaff.gif"" border=""0"" alt=""""></a>&nbsp;"
			ElseIf oRS.Collect("ViewStaffOtherFlag") = "1" Then
				'���Č��ŉ{���ς�
				sViewStaffHtml = "<a href=""javascript:void(0)"" onclick='window.open(""/staff/icon_navi.html"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=420,height=300"")'><img src=""/img/viewstaffother.gif"" border=""0"" alt=""""></a>&nbsp;"
			End If
			If oRS.Collect("OrderScoutFlag") = "1" Then
				'���[�����M�ς�
				sScoutHtml = "<a href=""javascript:void(0)"" onclick='window.open(""/staff/icon_navi.html"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=420,height=300"")'><img src=""/img/contact.gif"" border=""0"" alt=""""></a>&nbsp;"
			ElseIf oRS.Collect("CompanyScoutFlag") = "1" Then
				'���Č����[�����M�ς�
				sScoutHtml = "<a href=""javascript:void(0)"" onclick='window.open(""/staff/icon_navi.html"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=420,height=300"")'><img src=""/img/other_send.gif"" border=""0"" alt=""""></a>&nbsp;"
			End If
			If oRS.Collect("MailReceiveFlag") = "1" Then
				'���[����M�ς�
				sViewStaffHtml = "<a href=""javascript:void(0)"" onclick='window.open(""/staff/icon_navi.html"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=420,height=300"")'><img src=""/img/mailreceive.gif"" border=""0"" alt=""""></a>&nbsp;"
			End If
			'</�A�C�R���ݒ�>

			'<���t���,�ꊇ���[��>
			If vUserType = "company" Then
				Response.Write "<div class=""m0"" style=""float:left;width:20%;"">"
				If InStr(LCase(Request.ServerVariables("URL")), "jinzai") = 0 Then
					If oRS.Collect("LumpMailFlag") = "0" And oRS.Collect("PublicFlag") = "1" Then
						Response.Write "<input id=""lumpmail" & sStaffCode & """ class=""btn1"" type=""button"" value=""�ꊇҰٗ\��"" style=""width:120px;"" onclick=""open('/company/lumpmail/reg.asp?staffcode=" & sStaffCode & "&amp;ordercode=" & vOrderCode & "','lumpmail','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=420,height=300');"">"
					ElseIf oRS.Collect("LumpMailFlag") = "1" Then
						Response.Write "<input class=""btn1"" type=""button"" value=""�ꊇҰٗ\���"" disabled style=""width:120px;"">"
					Else
						Response.Write "<input class=""btn1"" type=""button"" value=""�ꊇҰٗ\��s��"" disabled style=""width:120px;"">"
					End If
				End If
				Response.Write "</div>"
			End If
			Response.Write "<p class=""m0"" style=""float:left;width:80%;text-align:right;font-size:10px;"">"
			If ChkStr(dbViewStaffDay) <> "" Then
				If dbViewStaffDay < rRS.Collect("UpdateDay") Then
					Response.Write "<span style=""color:#ff0000;"">�����E�҂������X�V���܂���</span>&nbsp;&nbsp;"
					Response.Write "�ڍ׉{�����F" & GetDateStr(dbViewStaffDay, "/") & "&nbsp;&nbsp;"
					Response.Write "�X�V���F" & GetDateStr(rRS.Collect("UpdateDay"), "/") & "&nbsp;&nbsp;"
				End If
			End If
			'Response.Write "�ŏI�A�N�Z�X���F" & GetStrLastAccessDay(rRS.Collect("LastAccessDay"))
			Response.Write "</p>"
			Response.Write "<div style=""clear:both;""></div>"
			'</���t���,�ꊇ���[��>

			Response.Write "<table class=""" & sClassName & " cw"" border=""0"">"
			'<�P�s��>
			Response.Write "<thead>"
			Response.Write "<tr>"
			Response.Write "<th colspan=""4"" style=""text-align:left;"">"
			'Response.Write oRS.Collect("StaffCode") & "�@"

			'If Len(oRS.Collect("PrefectureName") & oRS.Collect("City")) > 10 Then
			'	Response.Write Left(oRS.Collect("PrefectureName") & oRS.Collect("City"),10) & "..."
			'Else
				Response.Write "�@" & oRS.Collect("PrefectureName")
			'End If

			'Response.Write "�ݏZ�i" & Left(oRS.Collect("Age"),1) & "�΁^" & oRS.Collect("Sex") & "��" & "�j"
			Response.Write "�ݏZ�i" & Left(oRS.Collect("Age"),1) & "0��^" & oRS.Collect("Sex") & "��" & "�j"
			'<���E�҂̏�ԃt���O>
			Response.Write sNewHtml
			Response.Write sHotHtml
			Response.Write sWorkerAlarmHtml
			Response.Write sScoutHtml
			Response.Write sViewStaffHtml
			'Response.Write sLastAccess
			'</���E�҂̏�ԃt���O>
			Response.Write "</th>"
			Response.Write "</tr>"
			Response.Write "</thead>"
			'<�P�s��>
			Response.Write "<tbody>"

			'<�Q�s��>
			Response.Write "<tr>"
			Response.Write "<th align=""center"" rowspan=""" & wRow & """>"
			'Response.Write "<a href=""javascript:PersonDetail('" & oRS.Collect("StaffCode") & "');"">"
			Response.Write "<a href=""" & sStaffDetailURL & """ target=""_blank"">"
			Response.Write "<img src=""/img/shousai.gif"" border=""0"" alt=""""><br>"
			Response.Write "<b>�ڍ�</b>"
			Response.Write "</a>"
			Response.Write "</th>"
			Response.Write "<th colspan=""2"">"
			Response.Write "���݂̏�"
			Response.Write "</th>"
			Response.Write "<td>"
			Response.Write oRS.Collect("OperateClassWebName")
			Response.Write "</td>"
			Response.Write "</tr>"
			'</�Q�s��>

			'<�R�s��>
			If ChkStr(oRS.Collect("CareerJobType")) <> "" Then
				Response.Write "<tr>"
				Response.Write "<th colspan=""2"">�o���E��</th>"
				Response.Write "<td>"
				'Response.Write Replace(ChkStr(oRS.Collect("CareerJobType")), vbCrLf, "<br>")
				Response.Write Replace(RegExpReplace(oRS.Collect("CareerJobType"), "(\(\d).*?(�N\))"), vbCrLf, "<br>")
				Response.Write "</td>"
				Response.Write "</tr>"
			End If
			'</�R�s��>

			'<�S�s��>
			Response.Write "<tr>"
			Response.Write "<th rowspan=""3"">��]����</th>"
			Response.Write "<th>�E��</th>"
			Response.Write "<td>"
			Response.Write Replace(ChkStr(oRS.Collect("HopeJobType")), vbCrLf, "<br>")
			Response.Write "</td>"
			Response.Write "</tr>"
			'</�S�s��>

			'<�T�s��>
			Response.Write "<tr>"
			Response.Write "<th>�Ζ��n</th>"
			Response.Write "<td>"
			Response.Write Replace(ChkStr(oRS.Collect("HopeWorkingPlace")), vbCrLf, "<br>")
			Response.Write "</td>"
			Response.Write "</tr>"
			'</�T�s��>

			'<�U�s��>
			Response.Write "<tr>"
			Response.Write "<th>�ٗp�`��</th>"
			Response.Write "<td>"
			Response.Write Replace(ChkStr(oRS.Collect("HopeWorkingType")), vbCrLf, "<br>")
			Response.Write "</td>"
			Response.Write "</tr>"
			'</�U�s��>

			'<�V�s��>
			'If Trim(ChkStr(wSelfPR)) <> "" Then
			'	Response.Write "<tr>"
			'	Response.Write "<th colspan=""2"">����PR</th>"
			'	Response.Write "<td>"

			'	If Len(wSelfPR) > 100 Then
			'		Response.Write Left(wSelfPR,100) & "..."
			'	Else
			'		Response.Write wSelfPR
			'	End If

			'	Response.Write "</td>"
			'	Response.Write "</tr>"
			'End If
			'</�V�s��>

			Response.Write "<tr>"
			Response.Write "<td style=""padding:0px;border-width:0px;width:50px;""></td>"
			Response.Write "<td style=""padding:0px;border-width:0px;width:75px;""></td>"
			Response.Write "<td style=""padding:0px;border-width:0px;width:75px;""></td>"
			Response.Write "<td style=""padding:0px;border-width:0px;width:400px;""></td>"
			Response.Write "</tr>"
			Response.Write "</tbody>"
			Response.Write "</table>"
			Response.Write "<br>"
		End If
	End If

	Call RSClose(oRS)
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̊�{��񕔕����o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F�����ƃi�r/staff/person_list.asp
'�@�@�@�F�����ƃi�r/order/company_order.asp
'���@�l�F
'���@���F2007/04/12 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileBase(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim sRS
	Dim oRS2
	Dim oRS3
	Dim sErrer
	Dim flgQE

	Dim sStaffCode
	Dim sTableClass
	Dim sComment

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	sTableClass = "pattern9"
	If vUserType = "company" Or vUserType = "dispatch" Then sTableClass = "pattern8"

%>
	<table class="profileSmart smartBlock" style="display:none;">
		<thead>
        	<tr>
    			<th colspan="2">��{���</th>
    		</tr>
    	</thead>
        <tbody>
        	<tr>
            	<th colspan="2" class="promidasi">�ғ��敪</th>
            </tr>
            <tr>
                <td colspan="2">
                <%
                    Response.Write rRS.Collect("OperateClassWeb")
                    If rRS.Collect("HopeWorkStartDay") <> "" Then
                        Response.Write "(�Ζ��J�n�\����F" & GetDateStr(rRS.Collect("HopeWorkStartDay"), "/") & ")"
                    End If
                %>
                </td>
            </tr>
            <tr>
            	<th colspan="2" class="promidasi">���ʁE�N��</th>
            </tr>
            <tr>
            	<td colspan="2">
                <%
					Response.Write rRS.Collect("Sex")
					Response.Write "�E" & rRS.Collect("Age") & "��"
                %>
                </td>
            </tr>
            <tr>
            	<th colspan="2" class="promidasi">�Z���n</th>
            </tr>
            <tr>
            	<td colspan="2"><%= rRS.Collect("PrefectureName") & rRS.Collect("City") %></td>
            </tr>
            <tr>
            	<th colspan="2" class="promidasi">����PR</th>
            </tr>
            <tr>
            	<td colspan="2">
                <%
					Response.Write Replace(ChkStr(rRS.Collect("SelfPR")), vbCrLf, "<br>")
				%>
                </td>
            </tr>
            
                <%
					sComment = Array("")
					
					sSQL = "SELECT ResultId FROM P_MyNaviResult WITH(NOLOCK) WHERE StaffCode = '" & sStaffCode & "';"
					flgQE = QUERYEXE(dbconn, oRS2, sSQL, sError)
					If GetRSState(oRS2) = True Then
						sSQL  = "SELECT specialcomment FROM MyNavi_Result WITH(NOLOCK) WHERE id = '" & oRS2.Collect("resultid") & "';"
						flgQE = QUERYEXE(dbconn, oRS3, sSQL, sError)
						If GetRSState(oRS3) = True Then
				
							sSQL = "SELECT point1,point2,point3,point4,point5,point6 FROM P_MyNaviResult WITH(NOLOCK) WHERE StaffCode = '" & sStaffCode & "';"
							flgQE = QUERYEXE(dbconn, sRS, sSQL, sError)
							If GetRSState(sRS) = True Then
								%>
                                <tr>
                                    <th colspan="2" class="promidasi">�K�E�f�f�u���Ԃ�i�r�v����</th>
                                </tr>
                                <tr>
                                    <td colspan="2">
                                <%
								
								sComment = Split(Replace(oRS3.Collect("specialcomment"),"���Ȃ�","���̕�"),"<br>")
								Response.Write sComment(0)
								
								Response.Write "</td>"
								Response.Write "</tr>"
							End If
						End If
					End If
					
					'<�Ŋ�w />
					Call DspProfileNearbyStationSmart(rDB, rRS, vUserType, vUserID)
					
					If (G_USERTYPE = "company" Or G_USERTYPE = "dispatch" Or G_USERTYPE = "talent" Or G_USERTYPE = "staff") And Left(sStaffCode,1) = "T" Then
				%>
            
            <tr>
            	<th colspan="2" class="promidasi">�l���S�����E��</th>
			</tr>
            <tr>
            	<td colspan="2">
					<%= rRS.Collect("RecommendationLetter") %>
				</td>
            </tr>
			<% End If %>
            <% If ChkStr(rRS.Collect("Learn")) <> "" Then %>
            <tr>
            	<th colspan="2" class="promidasi">�w������ɐg�ɂ�������</th>
            </tr>
            <tr>
            	<td colspan="2"><%= rRS.Collect("Learn") %></td>
            </tr>
    		<% End If %>
    	</tbody>
    </table>
<%

	Response.Write "<table class=""" & sTableClass & " cw smartNone"" border=""0"" style=""margin-bottom:15px;"">"
	Response.Write "<colgroup>"
	Response.Write "<col style=""width:100px;"">"
	Response.Write "<col style=""width:100px;"">"
	Response.Write "<col style=""width:400px;"">"
	Response.Write "</colgroup>"
	Response.Write "<thead>"
	Response.Write "<tr>"
	Response.Write "<th colspan=""3"">��{���</th>"
	Response.Write "</tr>"
	Response.Write "</thead>"
	Response.Write "<tbody>"
	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">�ғ��敪</td>"
	Response.Write "<td>"
	Response.Write rRS.Collect("OperateClassWeb")
	If rRS.Collect("HopeWorkStartDay") <> "" Then
		Response.Write "(�Ζ��J�n�\����F" & GetDateStr(rRS.Collect("HopeWorkStartDay"), "/") & ")"
	End If
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">���ʁE�N��</th>"
	Response.Write "<td>"
	Response.Write rRS.Collect("Sex")
	Response.Write "�E" & rRS.Collect("Age") & "��"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">�Z���n</th>"
	Response.Write "<td>"
	Response.Write rRS.Collect("PrefectureName") & rRS.Collect("City")
	Response.Write "</td>"
	Response.Write "</tr>"

	'<���Ȃo�q>
	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">����PR</th>"
	Response.Write "<td>"
	Response.Write Replace(ChkStr(rRS.Collect("SelfPR")), vbCrLf, "<br>")
	Response.Write "</td>"
	Response.Write "</tr>"
	'</���Ȃo�q>

	'<���Ԃ�i�r����>
	sComment = Array("")

	sSQL = "SELECT ResultId FROM P_MyNaviResult WITH(NOLOCK) WHERE StaffCode = '" & sStaffCode & "';"
	flgQE = QUERYEXE(dbconn, oRS2, sSQL, sError)
	If GetRSState(oRS2) = True Then
		sSQL  = "SELECT specialcomment FROM MyNavi_Result WITH(NOLOCK) WHERE id = '" & oRS2.Collect("resultid") & "';"
		flgQE = QUERYEXE(dbconn, oRS3, sSQL, sError)
		If GetRSState(oRS3) = True Then

			sSQL = "SELECT point1,point2,point3,point4,point5,point6 FROM P_MyNaviResult WITH(NOLOCK) WHERE StaffCode = '" & sStaffCode & "';"
			flgQE = QUERYEXE(dbconn, sRS, sSQL, sError)
			If GetRSState(sRS) = True Then
				Response.Write "<tr>"
				Response.Write "<th colspan=""2"">�K�E�f�f�u���Ԃ�i�r�v����<br></th>"
				Response.Write "<td>"
				Response.Write "<div style=""width:120px; float:right; font-size:10px; color:#666699; height:100px;"" class=""smartFloat"">"
				Response.Write "<embed src=""/img/order/jinzai_radar.swf?point1=" & sRS.Collect("point1") & "&point2=" & sRS.Collect("point2") & "&point3=" & sRS.Collect("point3") & "&point4=" & sRS.Collect("point4") & "&point5=" & sRS.Collect("point5") & "&point6=" & sRS.Collect("point6") & """ quality=""high"" width=""120px"" height=""100px"" bgcolor=""#ffffff"" name=""mynavi_radar"" align=""middle"" menu=""false"">"
				Response.Write "</div>"
				sComment = Split(Replace(oRS3.Collect("specialcomment"),"���Ȃ�","���̕�"),"<br>")
				Response.Write sComment(0)
				Response.Write "</td>"
				Response.Write "</tr>"
			End If
		End If
	End If
	'</���Ԃ�i�r����>

	'<�Ŋ�w />
	Call DspProfileNearbyStation(rDB, rRS, vUserType, vUserID)

	'<�肷���[�Ɛl���S�����E��>
	If (G_USERTYPE = "company" Or G_USERTYPE = "dispatch" Or G_USERTYPE = "talent" Or G_USERTYPE = "staff") And Left(sStaffCode,1) = "T" Then
		Response.Write "<tr>"
		Response.Write "<th colspan=""2"">�l���S�����E��</th>"
		Response.Write "<td>"
		Response.Write rRS.Collect("RecommendationLetter")
		Response.Write "</td>"
		Response.Write "</tr>"
	End If
	'</�肷���[�Ɛl���S�����E��>

	'<�w�� />
	Call DspProfileEducateHistory(rDB, rRS, vUserType, vUserID)

	'<�g�ɂ�������>
	If ChkStr(rRS.Collect("Learn")) <> "" Then
		Response.Write "<tr>"
		Response.Write "<th colspan=""2"">�w������ɐg�ɂ�������</th>"
		Response.Write "<td>"
		Response.Write rRS.Collect("Learn")
		Response.Write "</td>"
		Response.Write "</tr>"
	End If
	'</�g�ɂ�������>

	Response.Write "</tbody>"
	Response.Write "</table>"
	Response.Write "<p style=""text-align:right;"" class=""m0""><a class=""stext"" href=""#pagetop"">���y�[�WTOP��</a></p>" & vbCrLf
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̍Ŋ�w�������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'���@���F2007/04/12 LIS K.Kokubo �쐬
'�@�@�@�F2008/10/28 LIS K.Kokubo up_GetDataNearbyStation�p�~��up_LstP_NearbyStation�ɕύX
'******************************************************************************
Function DspProfileNearbyStation(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sRailwayLine
	Dim sStation

	If GetRSState(rRS) = False Then Exit Function

	'�Ŋ�w
	sSQL = "EXEC up_LstP_NearbyStation '" & rRS.Collect("StaffCode") & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	sStation = ""
	Do While GetRSState(oRS) = True
		If sStation <> "" Then sStation = sStation & "<br>"
		sStation = sStation & oRS.Collect("StationName")
		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	'�Ŋ񉈐�
	sSQL = "EXEC up_LstP_NearbyRailwayLine '" & rRS.Collect("StaffCode") & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	sRailwayLine = ""
	Do While GetRSState(oRS) = True
		If sRailwayLine <> "" Then sRailwayLine = sRailwayLine & "<br>"
		sRailwayLine = sRailwayLine & oRS.Collect("RailwayCompanyName") & "�@" &oRS.Collect("RailwayLineName")
		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	Set oRS = Nothing

	Response.Write "<tr>"
	Response.Write "<th rowspan=""2"">�Ŋ�w</th>"
	Response.Write "<th>����</th>"
	Response.Write "<td>" & sRailwayLine & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th>�w</th>"
	Response.Write "<td>" & sStation & "</td>"
	Response.Write "</tr>"
End Function


'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̍Ŋ�w�������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'���@���F2007/04/12 LIS K.Kokubo �쐬
'�@�@�@�F2008/10/28 LIS K.Kokubo up_GetDataNearbyStation�p�~��up_LstP_NearbyStation�ɕύX
'******************************************************************************
Function DspProfileNearbyStationSmart(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sRailwayLine
	Dim sStation

	If GetRSState(rRS) = False Then Exit Function

	'�Ŋ�w
	sSQL = "EXEC up_LstP_NearbyStation '" & rRS.Collect("StaffCode") & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	sStation = ""
	Do While GetRSState(oRS) = True
		If sStation <> "" Then sStation = sStation & "<br>"
		sStation = sStation & oRS.Collect("StationName")
		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	'�Ŋ񉈐�
	sSQL = "EXEC up_LstP_NearbyRailwayLine '" & rRS.Collect("StaffCode") & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	sRailwayLine = ""
	Do While GetRSState(oRS) = True
		If sRailwayLine <> "" Then sRailwayLine = sRailwayLine & "<br>"
		sRailwayLine = sRailwayLine & oRS.Collect("RailwayCompanyName") & "�@" &oRS.Collect("RailwayLineName")
		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	Set oRS = Nothing

	Response.Write "<tr>"
	Response.Write "<th colspan=""2""  class=""promidasi"">�Ŋ�w</th>"
	response.write "</tr>"
	response.write "<tr>"
	Response.Write "<th>����</th>"
	Response.Write "<td>" & sRailwayLine & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th>�w</th>"
	Response.Write "<td>" & sStation & "</td>"
	Response.Write "</tr>"
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̊w����񕔕����o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/12 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileEducateHistory(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim idxEducateSchool

	sSQL = "sp_GetDataEducateHistory '" & rRS.Collect("StaffCode") & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

	If GetRSState(oRS) = False Then
		'�w���Ȃ�
		Response.Write "<tr>"
		Response.Write "<th colspan=""2"">�w��</th>"
		Response.Write "<td></td>"
		Response.Write "</tr>"
	Else
		'�w������
		idxEducateSchool = 1
		Do While GetRSState(oRS) = True
			If idxEducateSchool = 1 Then
				Response.Write "<tr>"
				Response.Write "<th rowspan=""" & oRS.RecordCount & """>�w��</th>"
			Else
				Response.Write "<tr>"
			End If
			Response.Write "<th>" & oRS.Collect("SchoolType") & "</th>"
			Response.Write "<td>"
			Response.Write oRS.Collect("Speciality")
			Response.Write "(" & Year(oRS.Collect("GraduateDay")) & "�N" & oRS.Collect("GraduateType") & ")"
			Response.Write "</td>"
			Response.Write "</tr>"

			oRS.MoveNext
			idxEducateSchool = idxEducateSchool + 1
		Loop
	End If
	Call RSClose(oRS)
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̐E����񕔕����o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'���@���F2007/04/12 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileCareerHistory(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sTableClass
	Dim sBusinessCareerMemo
	Dim iYearPeriod
	Dim iMonthPeriod
	Dim sEntryDay
	Dim sEntryYear
	Dim sEntryMonth
	Dim sRetireDay
	Dim sRetireYear
	Dim sRetireMonth
	Dim idx:	idx = 1

	If GetRSState(rRS) = False Then Exit Function

	sTableClass = "pattern9"
	If vUserType = "company" Or vUserType = "dispatch" Then sTableClass = "pattern8"

	sSQL = "sp_GetDataNote '" & rRS.Collect("StaffCode") & "', 'BusinessCareerMemo'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		sBusinessCareerMemo = oRS.Collect("Note")
	End If
	Call RSClose(oRS)

	sSQL = "sp_GetDataCareerHistory '" & rRS.Collect("StaffCode") & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

	If GetRSState(oRS) = True Then
		Response.Write "<table class=""" & sTableClass & " cw smartNone"" border=""0"" style=""margin-bottom:15px;"">"
		Response.Write "<colgroup>"
		Response.Write "<col style=""width:100px;"">"
		Response.Write "<col style=""width:100px;"">"
		Response.Write "<col style=""width:400px;"">"
		Response.Write "</colgroup>"
		Response.Write "<thead>"
		Response.Write "<tr>"
		Response.Write "<th colspan=""3"">�E��</th>"
		Response.Write "</tr>"
		Response.Write "</thead>"
		Response.Write "<tbody>"

		Do While GetRSState(oRS) = True
			sEntryDay = ""
			sEntryYear = ""
			sEntryMonth = ""
			sRetireDay = ""
			sRetireYear = ""
			sRetireMonth = ""

			'�A�Ɗ��Ԍv�Z
			iYearPeriod = Int(DateDiff("m", oRS.Collect("EntryDay"), oRS.Collect("RetireDay")) / 12)
			iMonthPeriod = DateDiff("m", oRS.Collect("EntryDay"), oRS.Collect("RetireDay")) Mod 12 + 1
			If iMonthPeriod = 12 Then
				iMonthPeriod = 0
				iYearPeriod = iYearPeriod + 1
			End If

			If IsDate(oRS.Collect("EntryDay")) = True Then
				sEntryYear = Year(oRS.Collect("EntryDay"))
				sEntryMonth = Month(oRS.Collect("EntryDay"))
				If sEntryYear & sEntryMonth <> "" Then sEntryDay = sEntryYear & "�N" & sEntryMonth & "��"
			End If
			If IsDate(oRS.Collect("RetireDay")) = True Then
				sRetireYear = Year(oRS.Collect("RetireDay"))
				sRetireMonth = Month(oRS.Collect("RetireDay"))
				If sRetireYear & sRetireMonth <> "" Then sRetireDay = sRetireYear & "�N" & sRetireMonth & "��"
			End If

			Response.Write "<tr>"
			Response.Write "<th rowspan=""5"">�E��" & idx & "</th>"
			Response.Write "<th>�Α�����</th>"
			Response.Write "<td>"

			If sEntryDay & sRetireDay <> "" Then Response.Write sEntryDay & "�`" & sRetireDay
			If ChkStr(iYearPeriod) & ChkStr(iMonthPeriod) <> "" Then
				Response.Write "&nbsp;("
				If iYearPeriod > 0 Then Response.Write iYearPeriod & "�N"
				If iMonthPeriod > 0 Then Response.Write iMonthPeriod & "����"
				Response.Write ")"
			End If

			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<th>�Ǝ�</th>"
			Response.Write "<td>"
			Response.Write oRS.Collect("IndustryType")
			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<th>�E��</th>"
			Response.Write "<td>"
			Response.Write oRS.Collect("JobType")
			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<th>�Ζ����e</th>"
			Response.Write "<td>"
			Response.Write Replace(ChkStr(oRS.Collect("BusinessDetail")), vbCrLf, "<br>")
			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<th>�Ζ��`��</th>"
			Response.Write "<td>"
			Response.Write oRS.Collect("WorkingType")
			Response.Write "</td>"
			Response.Write "</tr>"

			idx = idx + 1
			oRS.MoveNext
		Loop

		Response.Write "</tbody>"
		Response.Write "</table>"
		Response.Write "<p style=""text-align:right;"" class=""m0""><a class=""stext"" href=""#pagetop"">���y�[�WTOP��</a></p>"
	End If
	Call RSClose(oRS)
	
	If sBusinessCareerMemo <> "" Then
	
		'���̑��E��
		Response.Write "<table class=""" & sTableClass & " cw smartNone"" border=""0"" style=""margin-bottom:15px;"">"
		Response.Write "<colgroup>"
		Response.Write "<col style=""width:200px;"">"
		Response.Write "<col style=""width:400px;"">"
		Response.Write "</colgroup>"
		Response.Write "<thead>"
		Response.Write "<tr>"
		Response.Write "<th colspan=""2"">���̑�</th>"
		Response.Write "</tr>"
		Response.Write "</thead>"
		Response.Write "<tbody>"
		Response.Write "<tr>"
		Response.Write "<th>�E������</th>"
		Response.Write "<td>"
		Response.Write Replace(sBusinessCareerMemo,vbCrLf,"<br>")
		Response.Write "</td>"
		Response.Write "</tr>"
		Response.Write "</tbody>"
		Response.Write "</table>"
		Response.Write "<p style=""text-align:right;"" class=""m0""><a class=""stext"" href=""#pagetop"">���y�[�WTOP��</a></p>"
	End If

	Response.Write vbCrLf
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̐E����񕔕����o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'���@���F2007/04/12 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileCareerHistorySmart(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sTableClass
	Dim sBusinessCareerMemo
	Dim iYearPeriod
	Dim iMonthPeriod
	Dim sEntryDay
	Dim sEntryYear
	Dim sEntryMonth
	Dim sRetireDay
	Dim sRetireYear
	Dim sRetireMonth
	Dim idx:	idx = 1

	If GetRSState(rRS) = False Then Exit Function

	sTableClass = "pattern9"
	If vUserType = "company" Or vUserType = "dispatch" Then sTableClass = "pattern8"

	sSQL = "sp_GetDataNote '" & rRS.Collect("StaffCode") & "', 'BusinessCareerMemo'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		sBusinessCareerMemo = oRS.Collect("Note")
	End If
	Call RSClose(oRS)

	sSQL = "sp_GetDataCareerHistory '" & rRS.Collect("StaffCode") & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

	If GetRSState(oRS) = True Then
	
	%>
		<table class="profileSmart smartBlock" style="display:none;">
		<thead>
        	<tr>
    			<th colspan="2">�E��</th>
    		</tr>
    	</thead>
        <tbody>
        <%
		
		Do While GetRSState(oRS) = True
			sEntryDay = ""
			sEntryYear = ""
			sEntryMonth = ""
			sRetireDay = ""
			sRetireYear = ""
			sRetireMonth = ""

			'�A�Ɗ��Ԍv�Z
			iYearPeriod = Int(DateDiff("m", oRS.Collect("EntryDay"), oRS.Collect("RetireDay")) / 12)
			iMonthPeriod = DateDiff("m", oRS.Collect("EntryDay"), oRS.Collect("RetireDay")) Mod 12 + 1
			If iMonthPeriod = 12 Then
				iMonthPeriod = 0
				iYearPeriod = iYearPeriod + 1
			End If

			If IsDate(oRS.Collect("EntryDay")) = True Then
				sEntryYear = Year(oRS.Collect("EntryDay"))
				sEntryMonth = Month(oRS.Collect("EntryDay"))
				If sEntryYear & sEntryMonth <> "" Then sEntryDay = sEntryYear & "�N" & sEntryMonth & "��"
			End If
			If IsDate(oRS.Collect("RetireDay")) = True Then
				sRetireYear = Year(oRS.Collect("RetireDay"))
				sRetireMonth = Month(oRS.Collect("RetireDay"))
				If sRetireYear & sRetireMonth <> "" Then sRetireDay = sRetireYear & "�N" & sRetireMonth & "��"
			End If
		
		%>
			<tr>
    			<th colspan="2" class="promidasi">�E��<%= idx %></th>
            </tr>
            <tr>
            	<th>�Α�����</th>
                <td>
                <%
				
					If sEntryDay & sRetireDay <> "" Then Response.Write sEntryDay & "�`" & sRetireDay
					If ChkStr(iYearPeriod) & ChkStr(iMonthPeriod) <> "" Then
						Response.Write "&nbsp;("
						If iYearPeriod > 0 Then Response.Write iYearPeriod & "�N"
						If iMonthPeriod > 0 Then Response.Write iMonthPeriod & "����"
						Response.Write ")"
					End If
				
				%>
                </td>
            </tr>
            <tr>
            	<th>�Ǝ�</th>
                <td><%= oRS.Collect("IndustryType") %></td>
            </tr>
            <tr>
            	<th>�E��</th>
                <td><%= oRS.Collect("JobType") %></td>
            </tr>
            <tr>
            	<th>�Ζ����e</th>
                <td>
                <%
					Response.Write Replace(ChkStr(oRS.Collect("BusinessDetail")), vbCrLf, "<br>")
				%>
                </td>
            </tr>
            <tr>
            	<th>�Ζ��`��</th>
                <td><%= oRS.Collect("WorkingType") %></td>
            </tr>
            <%
			
				idx = idx + 1
				oRS.MoveNext
			Loop
			
  	        %>  
    	</tbody>
	</table>   
    <%

	End If
	Call RSClose(oRS)

	If sBusinessCareerMemo <> "" Then
	%>
		<table class="profileSmart smartBlock" style="display:none;">
            <thead>
                <tr>
                    <th colspan="2">���̑�</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <th colspan="2" class="promidasi">�E������</th>
                </tr>
                    <tr>
                    <td>
                    <%
                        Response.Write Replace(sBusinessCareerMemo,vbCrLf,"<br>")
                    %>
                    </td>
                </tr>
            </tbody>
        </table>            
	<%
	End If

	Response.Write vbCrLf
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̂h�s�E����񕔕����o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/12 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileCareerHistoryIT(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim oRS2
	Dim flgQE
	Dim sError

	Dim idx
	Dim sTempDay_it
	Dim sRangeDay_it,sTableClass

	Dim sOSLanguage	: sOSLanguage = ""
	Dim sDBTool		: sDBTool = ""

	If GetRSState(rRS) = False Then Exit Function

	sTableClass = "pattern9"
	If vUserType = "company" Or vUserType = "dispatch" Then sTableClass = "pattern8"

	sSQL = "sp_GetDataCareerHistoryIT '" & rRS.Collect("StaffCode") & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then

		Response.Write "<table class=""" & sTableClass & " cw smartNone"" border=""0"" style=""margin-bottom:15px;"">"
		Response.Write "<colgroup>"
		Response.Write "<col style=""width:100px;"">"
		Response.Write "<col style=""width:100px;"">"
		Response.Write "<col style=""width:400px;"">"
		Response.Write "</colgroup>"
		Response.Write "<thead>"
		Response.Write "<tr>"
		Response.Write "<th colspan=""3"">IT�n�E���ڍ�</th>"
		Response.Write "</tr>"
		Response.Write "</thead>"
		Response.Write "<tbody>"

		sSQL = "sp_GetDataDevelopmentTool '" & rRS.Collect("StaffCode") & "', '', ''"
		flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)

		idx = 1
		Do While GetRSState(oRS) = True
			sTempDay_it = ""
			sRangeDay_it = ""
			If oRS.Collect("StartDay") <> "" Then
				sTempDay_it = GetDateStr(oRS.Collect("StartDay"), "/")
				sTempDay_it = Year(sTempDay_it) & "/" & Month(sTempDay_it)
				sRangeDay_it = sTempDay_it & "�`"
			End If
			If oRS.Collect("EndDay") <> "" Then
				sTempDay_it = GetDateStr(oRS.Collect("EndDay"), "/")
				sTempDay_it = Year(sTempDay_it) & "/" & Month(sTempDay_it)
				If sRangeDay_it = "" Then sRangeDay_it = "�`"
				sRangeDay_it = sRangeDay_it & sTempDay_it
			End If

			Response.Write "<tr>"
			Response.Write "<th rowspan=""7"">IT�E��" & idx & "<img src=""/img/spacer.gif"" width=""1"" height=""1""></th>"
			Response.Write "<th>����</th>"
			Response.Write "<td>"
			Response.Write sRangeDay_it
			If oRS.Collect("StartDay") <> "" And oRS.Collect("EndDay") <> "" Then
				If Int(DateDiff("m",oRS.Collect("StartDay"),oRS.Collect("EndDay")) / 12) = 0 And DateDiff("m",oRS.Collect("StartDay"),oRS.Collect("EndDay")) mod 12 < 11 Then
					Response.Write "(" & DateDiff("m",oRS.Collect("StartDay"),oRS.Collect("EndDay")) mod 12 + 1 & "����)"
				ElseIf DateDiff("m",oRS.Collect("StartDay"),oRS.Collect("EndDay")) mod 12 = 11 Then
					Response.Write "(1�N)"
				Else
					Response.Write "(" & Int(DateDiff("m",oRS.Collect("StartDay"),oRS.Collect("EndDay")) / 12) & "�N" & DateDiff("m",oRS.Collect("StartDay"),oRS.Collect("EndDay")) mod 12 + 1 & "����)"
				End If
			End If
			Response.Write "</td>"
			Response.Write "</tr>"

			Response.Write "<tr>"
			Response.Write "<th>�J�����e</th>"
			Response.Write "<td>"
			Response.Write Replace(ChkStr(oRS.Collect("DevelopmentDetail")),vbCrLf,"<br>")
			Response.Write "</td>"
			Response.Write "</tr>"

			Response.Write "<tr>"
			Response.Write "<th>����</th>"
			Response.Write "<td>"
			Dim sType1 : sType1 = ""
			If oRS.Collect("PMFlag") = "1" Then sType1 = sType1 & "�@PM"
			If oRS.Collect("PLFlag") = "1" Then sType1 = sType1 & "�@PL"
			If oRS.Collect("SEFlag") = "1" Then sType1 = sType1 & "�@SE"
			If oRS.Collect("PGFlag") = "1" Then sType1 = sType1 & "�@PG"
			If oRS.Collect("TSFlag") = "1" Then sType1 = sType1 & "�@TS"
			If sType1 <> "" Then sType1 = Mid(sType1, 2)
			Response.Write sType1
			Response.Write "</td>"
			Response.Write "</tr>"

			Response.Write "<tr>"
			Response.Write "<th>��Ɠ��e</th>"
			Response.Write "<td>"

			Dim sType2 : sType2 = ""
			If oRS.Collect("SystemAnalysisFlag") = "1" Then sType2 = sType2 & "�@�V�X�e������"
			If oRS.Collect("DesignFlag") = "1" Then sType2 = sType2 & "�@�݌v"
			If oRS.Collect("DevelopmentFlag") = "1" Then sType2 = sType2 & "�@�J��"
			If oRS.Collect("TestFlag") = "1" Then sType2 = sType2 & "�@�e�X�g"
			If oRS.Collect("MaintenanceFlag") = "1" Then sType2 = sType2 & "�@�^�p�ێ�"
			If sType2 <> "" Then sType2 = Mid(sType2, 2)
			Response.Write sType2
			Response.Write "</td>"
			Response.Write "</tr>"

			Response.Write "<tr>"
			Response.Write "<th>�v���W�F�N�g�l��</th>"
			Response.Write "<td>"
			If oRS.Collect("Number") <> "" Then
				Response.Write oRS.Collect("Number") & "�l"
			End If
			Response.Write "</td>"
			Response.Write "</tr>"

			Response.Write "<tr>"
			Response.Write "<th>"
			sOSLanguage=""
			sDBTool = ""
			If oRS2.State <> 0 Then
				oRS2.Filter = "CareerHistoryITID = " & idx
				Do While GetRSState(oRS2) = True
					If oRS2.Collect("CategoryCode") = "OS" _
					Or oRS2.Collect("CategoryCode") = "DevelopmentLanguage" Then
						'�g�pOS�A����
						If sOSLanguage <> "" Then sOSLanguage = sOSLanguage & "<br>"
						sOSLanguage = sOSLanguage & oRS2.Collect("DevelopmentToolName")
					Else
						'DB�A���̑�
						If sDBTool <> "" Then sDBTool = sDBTool & "<br>"
						sDBTool = sDBTool & oRS2.Collect("DevelopmentToolName")
					End If
					oRS2.MoveNext
				Loop
				oRS2.Filter = 0
			End If
			Response.Write "�g�pOS�^����"
			Response.Write "</th>"
			Response.Write "<td>"
			Response.Write sOSLanguage
			Response.Write "</td>"
			Response.Write "</tr>"

			Response.Write "<tr>"
			Response.Write "<th>�g�p�c�[��<br>�^DB�^���̑�</th>"
			Response.Write "<td>"
			If ChkStr(oRS.Collect("DevelopmentRemark")) <> "" Then
				If sDBTool <> "" Then sDBTool = sDBTool & "<br>"
				sDBTool = sDBTool & Replace(ChkStr(oRS.Collect("DevelopmentRemark")),vbCrLf,"<br>")
			End If
			Response.Write sDBTool
			Response.Write "</td>"
			Response.Write "</tr>"

			idx = idx + 1
			oRS.MoveNext
		Loop

		Call RSClose(oRS2)
		Call RSClose(oRS)

		Response.Write "</tbody>"
		Response.Write "</table>"
		Response.Write "<p style=""text-align:right;"" class=""m0""><a class=""stext"" href=""#pagetop"">���y�[�WTOP��</a></p>" & vbCrLf
	End If
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̂h�s�E����񕔕����o��(smart)
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/12 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileCareerHistoryITsmart(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim oRS2
	Dim flgQE
	Dim sError

	Dim idx
	Dim sTempDay_it
	Dim sRangeDay_it,sTableClass

	Dim sOSLanguage	: sOSLanguage = ""
	Dim sDBTool		: sDBTool = ""

	If GetRSState(rRS) = False Then Exit Function

	sTableClass = "pattern9"
	If vUserType = "company" Or vUserType = "dispatch" Then sTableClass = "pattern8"

	sSQL = "sp_GetDataCareerHistoryIT '" & rRS.Collect("StaffCode") & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
	
	%>
		<table class="profileSmart smartBlock" style="display:none;">
            <thead>
                <tr>
                    <th colspan="2">IT�n�E���ڍ�</th>
                </tr>
            </thead>
            <tbody>
    		<%
    			sSQL = "sp_GetDataDevelopmentTool '" & rRS.Collect("StaffCode") & "', '', ''"
				flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
		
				idx = 1
				Do While GetRSState(oRS) = True
					sTempDay_it = ""
					sRangeDay_it = ""
					If oRS.Collect("StartDay") <> "" Then
						sTempDay_it = GetDateStr(oRS.Collect("StartDay"), "/")
						sTempDay_it = Year(sTempDay_it) & "/" & Month(sTempDay_it)
						sRangeDay_it = sTempDay_it & "�`"
					End If
					If oRS.Collect("EndDay") <> "" Then
						sTempDay_it = GetDateStr(oRS.Collect("EndDay"), "/")
						sTempDay_it = Year(sTempDay_it) & "/" & Month(sTempDay_it)
						If sRangeDay_it = "" Then sRangeDay_it = "�`"
						sRangeDay_it = sRangeDay_it & sTempDay_it
					End If
    		%>
    			<tr>
    				<th colspan="2" class="promidasi">IT�E��<%= idx %></th>
    			</tr>
                <tr>
                	<th>����</th>
                    <td>
                    <%
						If oRS.Collect("StartDay") <> "" And oRS.Collect("EndDay") <> "" Then
							If Int(DateDiff("m",oRS.Collect("StartDay"),oRS.Collect("EndDay")) / 12) = 0 And DateDiff("m",oRS.Collect("StartDay"),oRS.Collect("EndDay")) mod 12 < 11 Then
								Response.Write "(" & DateDiff("m",oRS.Collect("StartDay"),oRS.Collect("EndDay")) mod 12 + 1 & "����)"
							ElseIf DateDiff("m",oRS.Collect("StartDay"),oRS.Collect("EndDay")) mod 12 = 11 Then
								Response.Write "(1�N)"
							Else
								Response.Write "(" & Int(DateDiff("m",oRS.Collect("StartDay"),oRS.Collect("EndDay")) / 12) & "�N" & DateDiff("m",oRS.Collect("StartDay"),oRS.Collect("EndDay")) mod 12 + 1 & "����)"
							End If
						End If
					%>
                    </td>
                </tr>
                <tr>
                	<th colspan="2" class="itNaiyo">�J�����e</th>
                </tr>
                <tr>
                	<td colspan="2">
                    <%
						Response.Write Replace(ChkStr(oRS.Collect("DevelopmentDetail")),vbCrLf,"<br>")
					%>
                    </td>
                </tr>
                <tr>
                	<th>����</th>
                    <td>
                    <%
						Dim sType1 : sType1 = ""
						If oRS.Collect("PMFlag") = "1" Then sType1 = sType1 & "�@PM"
						If oRS.Collect("PLFlag") = "1" Then sType1 = sType1 & "�@PL"
						If oRS.Collect("SEFlag") = "1" Then sType1 = sType1 & "�@SE"
						If oRS.Collect("PGFlag") = "1" Then sType1 = sType1 & "�@PG"
						If oRS.Collect("TSFlag") = "1" Then sType1 = sType1 & "�@TS"
						If sType1 <> "" Then sType1 = Mid(sType1, 2)
						Response.Write sType1
					%>
                    </td>
                </tr>
                <tr>
                	<th colspan="2" class="itNaiyo">��Ɠ��e</th>
                </tr>
                <tr>
                    <td colspan="2">
                    <%
						Dim sType2 : sType2 = ""
						If oRS.Collect("SystemAnalysisFlag") = "1" Then sType2 = sType2 & "�@�V�X�e������"
						If oRS.Collect("DesignFlag") = "1" Then sType2 = sType2 & "�@�݌v"
						If oRS.Collect("DevelopmentFlag") = "1" Then sType2 = sType2 & "�@�J��"
						If oRS.Collect("TestFlag") = "1" Then sType2 = sType2 & "�@�e�X�g"
						If oRS.Collect("MaintenanceFlag") = "1" Then sType2 = sType2 & "�@�^�p�ێ�"
						If sType2 <> "" Then sType2 = Mid(sType2, 2)
						Response.Write sType2
					%>
                    </td>
                </tr>
                <tr>
                	<th colspan="2" class="itNaiyo">�v���W�F�N�g�l��</th>
               	</tr>
                <tr>
                	<td colspan="2">
                    <%
						If oRS.Collect("Number") <> "" Then
							Response.Write oRS.Collect("Number") & "�l"
						End If
					%>
                    </td>
                </tr>
                <tr>
                	<th colspan="2" class="itNaiyo">
                    <%
						sOSLanguage=""
						sDBTool = ""
						If oRS2.State <> 0 Then
							oRS2.Filter = "CareerHistoryITID = " & idx
							Do While GetRSState(oRS2) = True
								If oRS2.Collect("CategoryCode") = "OS" _
								Or oRS2.Collect("CategoryCode") = "DevelopmentLanguage" Then
									'�g�pOS�A����
									If sOSLanguage <> "" Then sOSLanguage = sOSLanguage & "<br>"
									sOSLanguage = sOSLanguage & oRS2.Collect("DevelopmentToolName")
								Else
									'DB�A���̑�
									If sDBTool <> "" Then sDBTool = sDBTool & "<br>"
									sDBTool = sDBTool & oRS2.Collect("DevelopmentToolName")
								End If
								oRS2.MoveNext
							Loop
							oRS2.Filter = 0
						End If
					%>
                    �g�pOS�^����
                    </th>
                </tr>
                <tr>
                    <td colspan="2">
                    	<%= sOSLanguage %>
                    </td>
                </tr>
                <tr>
                	<th colspan="2" class="itNaiyo">�g�p�c�[��/DB/���̑�</th>
                </tr>
                <tr>
                    <td colspan="2">
                    <%
						If ChkStr(oRS.Collect("DevelopmentRemark")) <> "" Then
							If sDBTool <> "" Then sDBTool = sDBTool & "<br>"
							sDBTool = sDBTool & Replace(ChkStr(oRS.Collect("DevelopmentRemark")),vbCrLf,"<br>")
						End If
						Response.Write sDBTool
					%>
                    </td>
                </tr>
                <%
				idx = idx + 1
					oRS.MoveNext
				Loop
		
				Call RSClose(oRS2)
				Call RSClose(oRS)
				
				%>                
            </tbody>
        </table>				
	<%
			
	End If
End Function


'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̃X�L���������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/12 LIS K.Kokubo �쐬
'�@�@�@�F2008/07/08 LIS K.Kokubo �\���p���i���Ή�
'�@�@�@�F2008/12/08 LIS K.Kokubo ��w�X�L���Ή�
'******************************************************************************
Function DspProfileSkill(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim oRS2
	Dim flgQE
	Dim sError

	Dim dbStaffCode
	Dim dbLanguageSeq
	Dim dbLanguageCode
	Dim dbLanguageName
	Dim dbOtherLanguage
	Dim dbLanguageActionLevelName1
	Dim dbLanguageActionLevelName2
	Dim dbLanguageActionLevelName3

	Dim sTableClass
	Dim flgLicenseData		: flgLicenseData = False
	Dim flgLanguageData		: flgLanguageData = False
	Dim flgSkillData		: flgSkillData = False
	Dim sLicense			: sLicense = ""
	Dim sOtherLicense		: sOtherLicense = ""
	Dim sLanguage			: sLanguage = ""
	Dim sOA					: sOA = ""
	Dim sOS					: sOS = ""
	Dim sApplication		: sApplication = ""
	Dim sDatabase			: sDatabase = ""
	Dim sDevelopmentLanguage: sDevelopmentLanguage = ""
	Dim iRowSkill			: iRowSkill = 0
	Dim iCol				: iCol = 2

	If GetRSState(rRS) = False Then Exit Function

	dbStaffCode = rRS.Collect("StaffCode")
	sTableClass = "pattern9"
	If vUserType = "company" Or vUserType = "dispatch" Then sTableClass = "pattern8"

	'<���i�擾>
	sSQL = "EXEC sp_GetDataLicense '" & rRS.Collect("StaffCode") & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then flgLicenseData = True
	Do While GetRSState(oRS) = True
		If sLicense <> "" Then sLicense = sLicense & "<br>"
		sLicense = sLicense & oRS.Collect("LicenseNameDsp")
		If oRS.Collect("LicenseNameDsp") <> oRS.Collect("LicenseName") Then sLicense = sLicense & "(" & oRS.Collect("LicenseName") & ")"
		If ChkStr(oRS.Collect("GetDay")) <> "" Then sLicense = sLicense & "[" & Year(GetDateStr(oRS.Collect("GetDay"), "/")) & "�N�擾]"

		oRS.MoveNext
	Loop
	Call RSClose(oRS)
	'</���i�擾>

	'<���̑����i�擾>
	sSQL = "EXEC sp_GetDataNote '" & dbStaffCode & "', 'OtherLicense';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		flgLicenseData = True
		If sLicense <> "" Then sOtherLicense = "<hr size=""1"">"
		sOtherLicense = sOtherLicense & vbCrLf & oRS.Collect("Note")
	End If
	Call RSClose(oRS)
	'</���̑����i�擾>

	'<��w�X�L���擾>
	sSQL = "EXEC up_LstP_Skill_Language '" & dbStaffCode & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		Set oRS.ActiveConnection = Nothing
		sSQL = "EXEC up_LstP_Skill_LanguageLevel '" & dbStaffCode & "','';"
		flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
		If GetRSState(oRS2) = True Then Set oRS2.ActiveConnection = Nothing
	End If

	Do While GetRSState(oRS)
		dbLanguageSeq = oRS.Collect("LanguageSeq")
		dbLanguageCode = oRS.Collect("LanguageCode")
		dbLanguageName = oRS.Collect("LanguageName")
		dbOtherLanguage = oRS.Collect("OtherLanguage")
		flgLanguageData = True

		dbLanguageActionLevelName1 = ""
		dbLanguageActionLevelName2 = ""
		dbLanguageActionLevelName3 = ""

		If GetRSState(oRS2) = True Then
			'��b���x��
			oRS2.Filter = "LanguageSeq = '" & dbLanguageSeq & "' And LanguageActionCode = '1'"
			If GetRSState(oRS2) = True Then
				dbLanguageActionLevelName1 = oRS2.Collect("LanguageActionLevelName")
			End If
			oRS2.Filter = 0
			'�ǉ����x��
			oRS2.Filter = "LanguageSeq = '" & dbLanguageSeq & "' And LanguageActionCode = '2'"
			If GetRSState(oRS2) = True Then
				dbLanguageActionLevelName2 = oRS2.Collect("LanguageActionLevelName")
			End If
			oRS2.Filter = 0
			'�앶���x��
			oRS2.Filter = "LanguageSeq = '" & dbLanguageSeq & "' And LanguageActionCode = '3'"
			If GetRSState(oRS2) = True Then
				dbLanguageActionLevelName3 = oRS2.Collect("LanguageActionLevelName")
			End If
			oRS2.Filter = 0
		End If

		sLanguage = sLanguage & "["
		If dbLanguageCode <> "999" Then
			sLanguage = sLanguage & dbLanguageName
		Else
			sLanguage = sLanguage & dbOtherLanguage
		End If
		sLanguage = sLanguage & "]"
		sLanguage = sLanguage & "<br>"
		If dbLanguageActionLevelName1 <> "" Then sLanguage = sLanguage & "��b�F" & dbLanguageActionLevelName1 & "<br>"
		If dbLanguageActionLevelName2 <> "" Then sLanguage = sLanguage & "�ǉ��F" & dbLanguageActionLevelName2 & "<br>"
		If dbLanguageActionLevelName3 <> "" Then sLanguage = sLanguage & "�앶�F" & dbLanguageActionLevelName3 & "<br>"

		oRS.MoveNext
		If GetRSState(oRS) = True Then sLanguage = sLanguage & "<div class=""line1""></div>"
	Loop
	Call RSClose(oRS)
	Call RSClose(oRS2)
	'</��w�X�L���擾>

	'<�X�L���擾>
	sSQL = "EXEC sp_GetDataSkill '" & dbStaffCode & "', '';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		flgSkillData = True

		'OS
		oRS.Filter = "CategoryCode = 'OS'"
		If GetRSState(oRS) = True Then iRowSkill = iRowSkill + 1
		Do While GetRSState(oRS) = True
			If sOS <> "" Then sOS = sOS & "<br>"
			sOS = sOS & oRS.Collect("SkillName")
			If oRS.Collect("Period") <> "" Then
				sOS = sOS & "(" & oRS.Collect("Period") & "�N)"
			End If
			oRS.MoveNext
		Loop
		oRS.Filter = 0
		'If sOS = "" Then sOS = "�n�r�o���Ȃ�"

		'Application
		oRS.Filter = "CategoryCode = 'Application'"
		If GetRSState(oRS) = True Then iRowSkill = iRowSkill + 1
		Do While GetRSState(oRS) = True
			If sApplication <> "" Then sApplication = sApplication & "<br>"
			sApplication = sApplication & oRS.Collect("SkillName")
			If oRS.Collect("Period") <> "" Then
				sApplication = sApplication & "(" & oRS.Collect("Period") & "�N)"
			End If
			oRS.MoveNext
		Loop
		oRS.Filter = 0
		'If sApplication = "" Then sApplication = "�A�v���P�[�V�����o���Ȃ�"

		'Database
		oRS.Filter = "CategoryCode = 'Database'"
		If GetRSState(oRS) = True Then iRowSkill = iRowSkill + 1
		Do While GetRSState(oRS) = True
			If sDatabase <> "" Then sDatabase = sDatabase & "<br>"
			sDatabase = sDatabase & oRS.Collect("SkillName")
			If oRS.Collect("Period") <> "" Then
				sDatabase = sDatabase & "(" & oRS.Collect("Period") & "�N)"
			End If
			oRS.MoveNext
		Loop
		oRS.Filter = 0
		'If sDatabase = "" Then sDatabase = "�f�[�^�x�[�X�o���Ȃ�"

		'DevelopmentLanguage
		oRS.Filter = "CategoryCode = 'DevelopmentLanguage'"
		If GetRSState(oRS) = True Then iRowSkill = iRowSkill + 1
		Do While GetRSState(oRS) = True
			If sDevelopmentLanguage <> "" Then sDevelopmentLanguage = sDevelopmentLanguage & "<br>"
			sDevelopmentLanguage = sDevelopmentLanguage & oRS.Collect("SkillName")
			If oRS.Collect("Period") <> "" Then
				sDevelopmentLanguage = sDevelopmentLanguage & "(" & oRS.Collect("Period") & "�N)"
			End If
			oRS.MoveNext
		Loop
		'If sDevelopmentLanguage = "" Then sDevelopmentLanguage = "�J������o���Ȃ�"
	End If

	Call RSClose(oRS)
	'</�X�L���擾>

	If flgSkillData = True Then iCol = iCol + 1
	If flgLicenseData = True Or flgLanguageData = True Or flgSkillData = True Then
		
		%>
		<table class="profileSmart smartBlock" style="display:none;">
            <thead>
                <tr>
                    <th colspan="2">���i�E�X�L��</th>
                </tr>
            </thead>
			<tbody>
            <% If flgLicenseData = True Then %>
            	<tr>
            		<th>���i</th>
            		<td>
                    <%= sLicense %>
                    <%= sOtherLicense %>
                    </td>
            	</tr>
            <% End If %>
            <% If flgLanguageData = True Then %>
            	<tr>
            		<th>��w</th>
            		<td>
                    <% sLanguage %>
                    </td>
            	</tr>
            <% End If %>
            <% If flgSkillData = True Then %>
            	<tr>
            		<th colspan="2" class="promidasi">�S�ۗL�X�L��</th>
                </tr>
                <tr>
                	<th>OS��</th>
                    <td><%= sOS %></td>
                </tr>
                <tr>
                	<th>�A�v���P�[�V������</th>
                	<td><%= sApplication %></td>
                </tr>
                <tr>
                	<th>�J�����ꖼ</th>
                	<td><%= sDevelopmentLanguage %></td>
                </tr>
                <tr>
                	<th>�f�[�^�x�[�X��</th>
                	<td><%= sDatabase %></td>
                </tr>
            <% End If %>
            </tbody>
		</table>
        <%
		
		Response.Write "<table class=""" & sTableClass & " cw smartNone"" border=""0"" style=""margin-bottom:15px;"">"
		Response.Write "<colgroup>"
		Response.Write "<col style=""width:100px;"">"
		Response.Write "<col style=""width:100px;"">"
		Response.Write "<col style=""width:400px;"">"
		Response.Write "</colgroup>"
		Response.Write "<thead>"
		Response.Write "<tr>"
		Response.Write "<th colspan=""" & iCol & """>���i�E�X�L��</th>"
		Response.Write "</tr>"
		Response.Write "</thead>"
		Response.Write "<tbody>"

		If flgLicenseData = True Then
			Response.Write "<tr>"
			Response.Write "<th colspan=""" & iCol - 1 & """>���i</th>"
			Response.Write "<td>"
			Response.Write sLicense
			Response.Write sOtherLicense
			Response.Write "</td>"
			Response.Write "</tr>"
		End If

		If flgLanguageData = True Then
			Response.Write "<tr>"
			Response.Write "<th colspan=""" & iCol - 1 & """>��w</th>"
			Response.Write "<td>"
			Response.Write sLanguage
			Response.Write "</td>"
			Response.Write "</tr>"
		End If

		If flgSkillData = True Then
			Response.Write "<tr>"
			Response.Write "<th rowspan=""4"">�S�ۗL<br>�X�L��</th>"
			Response.Write "<th>OS��</th>"
			Response.Write "<td>"
			Response.Write sOS
			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<th style=""font-size:10px;"">�A�v���P�[�V������</th>"
			Response.Write "<td>"
			Response.Write sApplication
			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<th>�J�����ꖼ</th>"
			Response.Write "<td>"
			Response.Write sDevelopmentLanguage
			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<th>�f�[�^�x�[�X��</th>"
			Response.Write "<td>"
			Response.Write sDatabase
			Response.Write "</td>"
			Response.Write "</tr>"
		End If

		Response.Write "</body>"
		Response.Write "</table>"
		Response.Write "<p style=""text-align:right;"" class=""m0""><a class=""stext"" href=""#pagetop"">���y�[�WTOP��</a></p>"
	End If
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̃X�L���������o��(�ȈՓo�^)
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2012/06/07 �^�j�U���쐬
'******************************************************************************
Function DspProfileSkillSimple(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim oRS2
	Dim flgQE
	Dim sError

	Dim dbStaffCode
	Dim dbLanguageSeq
	Dim dbLanguageCode
	Dim dbLanguageName
	Dim dbOtherLanguage
	Dim dbLanguageActionLevelName1
	Dim dbLanguageActionLevelName2
	Dim dbLanguageActionLevelName3

	Dim sTableClass
	Dim flgLicenseData		: flgLicenseData = False
	Dim flgLanguageData		: flgLanguageData = False
	Dim flgSkillData		: flgSkillData = False
	Dim sLicense			: sLicense = ""
	Dim sOtherLicense		: sOtherLicense = ""
	Dim sLanguage			: sLanguage = ""
	Dim sOA					: sOA = ""
	Dim sOS					: sOS = ""
	Dim sApplication		: sApplication = ""
	Dim sDatabase			: sDatabase = ""
	Dim sDevelopmentLanguage: sDevelopmentLanguage = ""
	Dim iRowSkill			: iRowSkill = 0
	Dim iCol				: iCol = 2

	If GetRSState(rRS) = False Then Exit Function

	dbStaffCode = rRS.Collect("StaffCode")
	sTableClass = "pattern9"
	If vUserType = "company" Or vUserType = "dispatch" Then sTableClass = "pattern8"

	'<���i�擾>
	sSQL = "EXEC sp_GetDataLicense_Simple '" & rRS.Collect("StaffCode") & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then flgLicenseData = True
	Do While GetRSState(oRS) = True
		If sLicense <> "" Then sLicense = sLicense & "<br>"
		sLicense = sLicense & oRS.Collect("LicenseNameDsp")
		If oRS.Collect("LicenseNameDsp") <> oRS.Collect("LicenseName") Then sLicense = sLicense & "(" & oRS.Collect("LicenseName") & ")"
		If ChkStr(oRS.Collect("GetDay")) <> "" Then sLicense = sLicense & "[" & Year(GetDateStr(oRS.Collect("GetDay"), "/")) & "�N�擾]"

		oRS.MoveNext
	Loop
	Call RSClose(oRS)
	'</���i�擾>

	'<���̑����i�擾>
	sSQL = "EXEC sp_GetDataNote '" & dbStaffCode & "', 'OtherLicense';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		flgLicenseData = True
		If sLicense <> "" Then sOtherLicense = "<hr size=""1"">"
		sOtherLicense = sOtherLicense & vbCrLf & oRS.Collect("Note")
	End If
	Call RSClose(oRS)
	'</���̑����i�擾>

	'<��w�X�L���擾>
	sSQL = "EXEC up_LstP_Skill_Language '" & dbStaffCode & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		Set oRS.ActiveConnection = Nothing
		sSQL = "EXEC up_LstP_Skill_LanguageLevel '" & dbStaffCode & "','';"
		flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
		If GetRSState(oRS2) = True Then Set oRS2.ActiveConnection = Nothing
	End If

	Do While GetRSState(oRS)
		dbLanguageSeq = oRS.Collect("LanguageSeq")
		dbLanguageCode = oRS.Collect("LanguageCode")
		dbLanguageName = oRS.Collect("LanguageName")
		dbOtherLanguage = oRS.Collect("OtherLanguage")
		flgLanguageData = True

		dbLanguageActionLevelName1 = ""
		dbLanguageActionLevelName2 = ""
		dbLanguageActionLevelName3 = ""

		If GetRSState(oRS2) = True Then
			'��b���x��
			oRS2.Filter = "LanguageSeq = '" & dbLanguageSeq & "' And LanguageActionCode = '1'"
			If GetRSState(oRS2) = True Then
				dbLanguageActionLevelName1 = oRS2.Collect("LanguageActionLevelName")
			End If
			oRS2.Filter = 0
			'�ǉ����x��
			oRS2.Filter = "LanguageSeq = '" & dbLanguageSeq & "' And LanguageActionCode = '2'"
			If GetRSState(oRS2) = True Then
				dbLanguageActionLevelName2 = oRS2.Collect("LanguageActionLevelName")
			End If
			oRS2.Filter = 0
			'�앶���x��
			oRS2.Filter = "LanguageSeq = '" & dbLanguageSeq & "' And LanguageActionCode = '3'"
			If GetRSState(oRS2) = True Then
				dbLanguageActionLevelName3 = oRS2.Collect("LanguageActionLevelName")
			End If
			oRS2.Filter = 0
		End If

		sLanguage = sLanguage & "["
		If dbLanguageCode <> "999" Then
			sLanguage = sLanguage & dbLanguageName
		Else
			sLanguage = sLanguage & dbOtherLanguage
		End If
		sLanguage = sLanguage & "]"
		sLanguage = sLanguage & "<br>"
		If dbLanguageActionLevelName1 <> "" Then sLanguage = sLanguage & "��b�F" & dbLanguageActionLevelName1 & "<br>"
		If dbLanguageActionLevelName2 <> "" Then sLanguage = sLanguage & "�ǉ��F" & dbLanguageActionLevelName2 & "<br>"
		If dbLanguageActionLevelName3 <> "" Then sLanguage = sLanguage & "�앶�F" & dbLanguageActionLevelName3 & "<br>"

		oRS.MoveNext
		If GetRSState(oRS) = True Then sLanguage = sLanguage & "<div class=""line1""></div>"
	Loop
	Call RSClose(oRS)
	Call RSClose(oRS2)
	'</��w�X�L���擾>

	'<�X�L���擾>
	sSQL = "EXEC sp_GetDataSkill_Simple '" & dbStaffCode & "', '';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		flgSkillData = True

		'OS
		oRS.Filter = "CategoryCode = 'OS'"
		If GetRSState(oRS) = True Then iRowSkill = iRowSkill + 1
		Do While GetRSState(oRS) = True
			If sOS <> "" Then sOS = sOS & "<br>"
			sOS = sOS & oRS.Collect("SkillName")
			If oRS.Collect("Period") <> "" Then
				sOS = sOS & "(" & oRS.Collect("Period") & "�N)"
			End If
			oRS.MoveNext
		Loop
		oRS.Filter = 0
		'If sOS = "" Then sOS = "�n�r�o���Ȃ�"

		'Application
		oRS.Filter = "CategoryCode = 'Application'"
		If GetRSState(oRS) = True Then iRowSkill = iRowSkill + 1
		Do While GetRSState(oRS) = True
			If sApplication <> "" Then sApplication = sApplication & "<br>"
			sApplication = sApplication & oRS.Collect("SkillName")
			If oRS.Collect("Period") <> "" Then
				sApplication = sApplication & "(" & oRS.Collect("Period") & "�N)"
			End If
			oRS.MoveNext
		Loop
		oRS.Filter = 0
		'If sApplication = "" Then sApplication = "�A�v���P�[�V�����o���Ȃ�"

		'Database
		oRS.Filter = "CategoryCode = 'Database'"
		If GetRSState(oRS) = True Then iRowSkill = iRowSkill + 1
		Do While GetRSState(oRS) = True
			If sDatabase <> "" Then sDatabase = sDatabase & "<br>"
			sDatabase = sDatabase & oRS.Collect("SkillName")
			If oRS.Collect("Period") <> "" Then
				sDatabase = sDatabase & "(" & oRS.Collect("Period") & "�N)"
			End If
			oRS.MoveNext
		Loop
		oRS.Filter = 0
		'If sDatabase = "" Then sDatabase = "�f�[�^�x�[�X�o���Ȃ�"

		'DevelopmentLanguage
		oRS.Filter = "CategoryCode = 'DevelopmentLanguage'"
		If GetRSState(oRS) = True Then iRowSkill = iRowSkill + 1
		Do While GetRSState(oRS) = True
			If sDevelopmentLanguage <> "" Then sDevelopmentLanguage = sDevelopmentLanguage & "<br>"
			sDevelopmentLanguage = sDevelopmentLanguage & oRS.Collect("SkillName")
			If oRS.Collect("Period") <> "" Then
				sDevelopmentLanguage = sDevelopmentLanguage & "(" & oRS.Collect("Period") & "�N)"
			End If
			oRS.MoveNext
		Loop
		'If sDevelopmentLanguage = "" Then sDevelopmentLanguage = "�J������o���Ȃ�"
	End If

	Call RSClose(oRS)
	'</�X�L���擾>

	If flgSkillData = True Then iCol = iCol + 1
	If flgLicenseData = True Or flgLanguageData = True Or flgSkillData = True Then
		%>
		<table class="profileSmart smartBlock" style="display:none;">
            <thead>
                <tr>
                    <th colspan="2">���i�E�X�L��</th>
                </tr>
            </thead>
			<tbody>
            <% If flgLicenseData = True Then %>
            	<tr>
            		<th>���i</th>
            		<td>
                    <%= sLicense %>
                    <%= sOtherLicense %>
                    </td>
            	</tr>
            <% End If %>
            <% If flgLanguageData = True Then %>
            	<tr>
            		<th>��w</th>
            		<td>
                    <%= sLanguage %>
                    </td>
            	</tr>
            <% End If %>
            <% If flgSkillData = True Then %>
            	<tr>
            		<th colspan="2" class="promidasi">�S�ۗL�X�L��</th>
                </tr>
                <tr>
                	<th>OS��</th>
                    <td><%= sOS %></td>
                </tr>
                <tr>
                	<th>�A�v���P�[�V������</th>
                	<td><%= sApplication %></td>
                </tr>
                <tr>
                	<th>�J�����ꖼ</th>
                	<td><%= sDevelopmentLanguage %></td>
                </tr>
                <tr>
                	<th>�f�[�^�x�[�X��</th>
                	<td><%= sDatabase %></td>
                </tr>
            <% End If %>
            </tbody>
		</table>
        <%
	
	
		Response.Write "<table class=""" & sTableClass & " cw smartNone"" border=""0"" style=""margin-bottom:15px;"">"
		Response.Write "<colgroup>"
		Response.Write "<col style=""width:100px;"">"
		Response.Write "<col style=""width:100px;"">"
		Response.Write "<col style=""width:400px;"">"
		Response.Write "</colgroup>"
		Response.Write "<thead>"
		Response.Write "<tr>"
		Response.Write "<th colspan=""" & iCol & """>���i�E�X�L��</th>"
		Response.Write "</tr>"
		Response.Write "</thead>"
		Response.Write "<tbody>"

		If flgLicenseData = True Then
			Response.Write "<tr>"
			Response.Write "<th colspan=""" & iCol - 1 & """>���i</th>"
			Response.Write "<td>"
			Response.Write sLicense
			Response.Write sOtherLicense
			Response.Write "</td>"
			Response.Write "</tr>"
		End If

		If flgLanguageData = True Then
			Response.Write "<tr>"
			Response.Write "<th colspan=""" & iCol - 1 & """>��w</th>"
			Response.Write "<td>"
			Response.Write sLanguage
			Response.Write "</td>"
			Response.Write "</tr>"
		End If

		If flgSkillData = True Then
			Response.Write "<tr>"
			Response.Write "<th rowspan=""4"">�S�ۗL<br>�X�L��</th>"
			Response.Write "<th>OS��</th>"
			Response.Write "<td>"
			Response.Write sOS
			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<th style=""font-size:10px;"">�A�v���P�[�V������</th>"
			Response.Write "<td>"
			Response.Write sApplication
			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<th>�J�����ꖼ</th>"
			Response.Write "<td>"
			Response.Write sDevelopmentLanguage
			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<th>�f�[�^�x�[�X��</th>"
			Response.Write "<td>"
			Response.Write sDatabase
			Response.Write "</td>"
			Response.Write "</tr>"
		End If

		Response.Write "</body>"
		Response.Write "</table>"
		Response.Write "<p style=""text-align:right;"" class=""m0""><a class=""stext"" href=""#pagetop"">���y�[�WTOP��</a></p>"
	End If
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̊�]�����������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/12 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileHope(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vOrderCode)
	Dim sPattern

	If GetRSState(rRS) = False Then Exit Function

	If vUserType = "staff" Then
		sPattern = "pattern9"
	ElseIf vUserType = "company" Then
		sPattern = "pattern8"
	End If

	Response.Write "<table class=""" & sPattern & " cw smartNone"" border=""0"" style=""margin-bottom:15px;"">"
	Response.Write "<colgroup>"
	Response.Write "<col style=""width:100px;"">"
	Response.Write "<col style=""width:100px;"">"
	Response.Write "<col style=""width:400px;"">"
	Response.Write "</colgroup>"
	Response.Write "<thead>"
	Response.Write "<tr>"
	Response.Write "<th colspan=""3"">��]����</th>"
	Response.Write "</tr>"
	Response.Write "</thead>"
	Response.Write "<tbody>"

	'��]�Ζ��`��
	Call DspProfileHopeWorkingType(rDB, rRS, vUserType, vUserID)
    '��]�Ζ��`��
	Call DspProfileHopeContactableTime(rDB, rRS, vUserType, vUserID)
	'��]�Ǝ�
	Call DspProfileHopeIndustry(rDB, rRS, vUserType, vUserID)
	'��]�E��
	Call DspProfileHopeJobType(rDB, rRS, vUserType, vUserID)
    '����
    Call DspProfileSideJob(rDB, rRS, vUserType, vUserID)
	'��]�Ζ��n
	Call DspProfileHopeWorkingPlace(rDB, rRS, vUserType, vUserID)
	'���^����
	Call DspProfileHopeSalary(rDB, rRS, vUserType, vUserID)
	'���ԁE����
	Call DspProfileHopeSpan(rDB, rRS, vUserType, vUserID)
	'��������
	Call DspProfileHopeWelfare(rDB, rRS, vUserType, vUserID)

	Response.Write "</tbody>"
	Response.Write "</table>"
	Response.Write "<p style=""text-align:right;"" class=""m0""><a class=""stext"" href=""#pagetop"">���y�[�WTOP��</a></p>" & vbCrLf
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̊�]�����Ζ��`�ԕ������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/12 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileHopeWorkingType(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim sStaffCode
	Dim sWorkingType	: sWorkingType = ""

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	sSQL = "sp_GetDataWorkingType '" & sStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		sWorkingType = sWorkingType & oRS.Collect("WorkingTypeName")
		oRS.MoveNext
		If GetRSState(oRS) = True Then sWorkingType = sWorkingType & "�@"
	Loop
	Call RSClose(oRS)

	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">�Ζ��`��</th>"
	Response.Write "<td>"
	Response.Write sWorkingType
	Response.Write "</td>"
	Response.Write "</tr>"

	If G_USERTYPE = "dispatch" Then
		Response.Write "<tr>"
		Response.Write "<th colspan=""2"">���Ј��Љ�\��h���̊�]</th>"
		Response.Write "<td>"

		If rRS.Collect("TempToPermFlag") = "1" Then
			Response.Write "��]����"
		Else
			Response.Write "��]���Ȃ�"
		End If

		Response.Write "</td>"
		Response.Write "</tr>"
	End If
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̊�]�����Ǝ핔�����o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/12 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileHopeIndustry(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sStaffCode

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">�Ǝ�</th>"
	Response.Write "<td>"

	sSQL = "sp_GetDataIndustryType '" & sStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		Response.Write oRS.Collect("IndustryTypeName")
		oRS.MoveNext
		If GetRSState(oRS) = True Then Response.Write "<br>"
	Loop
	Call RSClose(oRS)

	Response.Write "</td>"
	Response.Write "</tr>"
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̊�]�����Ǝ핔�����o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/12 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileHopeJobType(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sStaffCode

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">�E��</th>"
	Response.Write "<td>"

	sSQL = "sp_GetDataHopeJobType '" & sStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		Response.Write oRS.Collect("JobTypeName")
		oRS.MoveNext
		If GetRSState(oRS) = True Then Response.Write "<br>"
	Loop
	Call RSClose(oRS)

	Response.Write "</td>"
	Response.Write "</tr>"
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̊�]�����Ζ��n�������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/12 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileHopeWorkingPlace(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sStaffCode

	Dim flgCommuteTime	: flgCommuteTime = False
	Dim flgStation		: flgStation = False
	Dim sArea			: sArea = ""
	Dim sPlace			: sPlace = ""
	Dim sHopeCommuteTime: sHopeCommuteTime = ""
	Dim sStation		: sStation = ""
	Dim sRailwayLine	: sRailwayLine = ""
	Dim iRow			: iRow = 1

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	sSQL = "sp_GetDataHopeWorkingPlace '" & sStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then iRow = iRow + 1
	Do While GetRSState(oRS) = True
		If InStr(sArea, oRS.Collect("Area")) = 0 Then
			'�d�����Ȃ��G���A���̂�
			sArea = sArea & "�@" & oRS.Collect("Area")
		End If
		sPlace = sPlace & oRS.Collect("WorkingPlace")
		oRS.MoveNext
		If GetRSState(oRS) = True Then sPlace = sPlace & "<br>"
	Loop
	Call RSClose(oRS)

	'��]�w
	sSQL = "sp_GetDataHopeCommuting '" & sStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		If sHopeCommuteTime <> "" Then sHopeCommuteTime = sHopeCommuteTime & "<br>"
		If ChkStr(oRS.Collect("HopeCommuteTime")) <> "" Then
			flgCommuteTime = True
			sHopeCommuteTime = oRS.Collect("HopeCommuteTime") & "��"
		End If

		If ChkStr(oRS.Collect("StationName")) <> "" Then
			flgStation = True
			If sStation <> "" Then sStation = sStation & "<br>"
			sStation = sStation & oRS.Collect("StationName") & "�w"

			If oRS.Collect("MinuteToStation") <> "" Then
				sStation = sStation & "(�w����" & oRS.Collect("MinuteToStation") & "���ȓ�)"
			End If
		End If
		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	If flgCommuteTime = True Then iRow = iRow + 1
	If flgStation = True Then iRow = iRow + 1

	'��]����
'	sSQL = "sp_GetDataHopeRailwayLine '" & sStaffCode & "'"
'	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
'	If GetRSState(oRS) = True Then iRow = iRow + 1
'	Do While GetRSState(oRS) = True
'		sRailwayLine = sRailwayLine & _
'			oRS.Collect("RailwayCompanyName") & _
'			"�@" & oRS.Collect("RailwayLineName") & "<br>"
'		oRS.MoveNext
'		If GetRSState(oRS) = True Then sRailwayLine = sRailwayLine & "<br>"
'	Loop
'	Call RSClose(oRS)

	Response.Write "<tr>"
	Response.Write "<th rowspan=""" & iRow & """>�Ζ��n</th>"
	Response.Write "<th>��]��</th>"
	Response.Write "<td>" & sArea & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th>��]�Ζ��n</th>"
	Response.Write "<td>" & sPlace & "</td>"
	Response.Write "</tr>"

	If sHopeCommuteTime <> "" Then
		Response.Write "<tr>"
		Response.Write "<th>��]�ʋΎ���</th>"
		Response.Write "<td>" & sHopeCommuteTime & "</td>"
		Response.Write "</tr>"
	End If

	If sStation <> "" Then
		Response.Write "<tr>"
		Response.Write "<th>��]�w</th>"
		Response.Write "<td>" & sStation & "</td>"
		Response.Write "</tr>"
	End If
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̊�]�������^�������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/12 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileHopeSalary(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sStaffCode
	Dim YMin
	Dim YMax
	Dim MMin
	Dim MMax
	Dim DMin
	Dim DMax
	Dim HMin
	Dim HMax
	Dim PercentagePay
	Dim Remark
	Dim iRow

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	If ChkStr(rRS.Collect("YearlyIncomeMin")) <> "" Then YMin = GetJapaneseYen(rRS.Collect("YearlyIncomeMin"))
	If ChkStr(rRS.Collect("YearlyIncomeMax")) <> "" Then YMax = GetJapaneseYen(rRS.Collect("YearlyIncomeMax"))
	If ChkStr(rRS.Collect("MonthlyIncomeMin")) <> "" Then MMin = GetJapaneseYen(rRS.Collect("MonthlyIncomeMin"))
	If ChkStr(rRS.Collect("MonthlyIncomeMax")) <> "" Then MMax = GetJapaneseYen(rRS.Collect("MonthlyIncomeMax"))
	If ChkStr(rRS.Collect("DailyIncomeMin")) <> "" Then DMin = GetJapaneseYen(rRS.Collect("DailyIncomeMin"))
	If ChkStr(rRS.Collect("DailyIncomeMax")) <> "" Then DMax = GetJapaneseYen(rRS.Collect("DailyIncomeMax"))
	If ChkStr(rRS.Collect("HourlyIncomeMin")) <> "" Then HMin = GetJapaneseYen(rRS.Collect("HourlyIncomeMin"))
	If ChkStr(rRS.Collect("HourlyIncomeMax")) <> "" Then HMax = GetJapaneseYen(rRS.Collect("HourlyIncomeMax"))
	PercentagePay = ChkStr(rRS.Collect("PercentagePayFlag"))
	Remark = ChkStr(rRS.Collect("IncomeRemark"))

	iRow = 0
	If YMin & YMax <> "" Then iRow = iRow + 1
	If MMin & MMax <> "" Then iRow = iRow + 1
	If DMin & DMax <> "" Then iRow = iRow + 1
	If HMin & HMax <> "" Then iRow = iRow + 1
	If PercentagePay <> "" Then iRow = iRow + 1
	If Remark <> "" Then iRow = iRow + 1

	If iRow > 0 Then
		Response.Write "<tr>"
		Response.Write "<th rowspan=""" & iRow & """>��]���^</th>"
		If YMin & YMax <> "" Then
			Response.Write "<th>�N��</th>"
			Response.Write "<td>" & YMin & "�`" & YMax &"</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
		End If

		If MMin & MMax <> "" Then
			Response.Write "<th>����</th>"
			Response.Write "<td>" & MMin & "�`" & MMax & "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
		End If

		If DMin & DMax <> "" Then
			Response.Write "<th>����</th>"
			Response.Write "<td>" & DMin & "�`" & DMax & "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
		End If

		If HMin & HMax <> "" Then
			Response.Write "<th>����</th>"
			Response.Write "<td>" & HMin & "�`" & HMax & "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
		End If

		Response.Write "<th>����</th>"
		Response.Write "<td>"
		If PercentagePay = "1" Then
			Response.Write "��]����"
		ElseIf PercentagePay = "0" Then
			Response.Write "��]���Ȃ�"
		Else
			Response.Write "�������Ȃ�"
		End If
		Response.Write "</td>"
		Response.Write "</tr>"

		If Remark <> "" Then
			Response.Write "<tr>"
			Response.Write "<th>���l</th>"
			Response.Write "<td>" & Remark & "</td>"
			Response.Write "</tr>"
		End If
	Else
		'��]���^�����������ꍇ
		Response.Write "<tr>"
		Response.Write "<th colspan=""2"">��]���^</th>"
		Response.Write "<td></td>"
		Response.Write "</tr>"
	End If
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̊�]�������ԁE���ԕ������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/12 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileHopeSpan(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sStaffCode
	Dim sTransfer
	Dim sWorkPeriod
	Dim sWorkPeriodFlag
	Dim sWorkMonthPeriod
	Dim sWorkTime
	Dim sWSTime
	Dim sWETime
	Dim sOverWork
	Dim sOverWorkTimeMax
	Dim sWorkShift
	Dim sHoliday

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	sTransfer = ChkStr(rRS.Collect("TransferFlag"))
	Select Case sTransfer
		Case "1":	sTransfer = "��"
		Case "0":	sTransfer = "�s��"
		Case Else:	sTransfer = "�������Ȃ�"
	End Select

	sWorkPeriod = ChkStr(oRS.Collect("WorkPeriod"))
	sWorkPeriodFlag = ChkStr(oRS.Collect("WorkPeriodFlag"))
	sWorkMonthPeriod = ChkStr(oRS.Collect("WorkMonthPeriod"))
	sWSTime = ChkStr(oRS.Collect("WorkStartTime"))
	sWETime = ChkStr(oRS.Collect("WorkEndTime"))
	If sWSTime & sWETime <> "" Then
		If sWSTime <> "" Then sWorkTime = sWorkTime & Left(sWSTime, 2) & ":" & Right(sWSTime, 2)
		sWorkTime = sWorkTime & "�`"
		If sWETime <> "" Then sWorkTime = sWorkTime & Left(sWETime, 2) & ":" & Right(sWETime, 2)
	End If

	sOverWork = ChkStr(rRS.Collect("OverWorkFlag"))
	Select Case sOverWork
		Case "1":	sOverWork = "��"
		Case "0":	sOverWork = "�s��"
		Case Else:	sOverWork = "�������Ȃ�"
	End Select
	sOverWorkTimeMax = ChkStr(rRS.Collect("OverWorkTimeMax"))
	If sOverWorkTimeMax <> "" Then sOverWorkTimeMax = Left(sOverWorkTimeMax, 2) & ":" & Right(sOverWorkTimeMax, 2)

	sWorkShift = ChkStr(rRS.Collect("WorkShiftFlag"))
	Select Case sWorkShift
		Case "1":	sWorkShift = "��"
		Case "0":	sWorkShift = "�s��"
		Case Else:	sWorkShift = "�������Ȃ�"
	End Select

	sHoliday = ""
	If oRS.Collect("MonHolidayFlag") = "1" Then sHoliday = sHoliday & "��"
	If oRS.Collect("TueHolidayFlag") = "1" Then sHoliday = sHoliday & "��"
	If oRS.Collect("WedHolidayFlag") = "1" Then sHoliday = sHoliday & "��"
	If oRS.Collect("ThuHolidayFlag") = "1" Then sHoliday = sHoliday & "��"
	If oRS.Collect("FriHolidayFlag") = "1" Then sHoliday = sHoliday & "��"
	If oRS.Collect("SatHolidayFlag") = "1" Then sHoliday = sHoliday & "�y"
	If oRS.Collect("SunHolidayFlag") = "1" Then sHoliday = sHoliday & "��"
	If oRS.Collect("PublicHolidayFlag") = "1" Then sHoliday = sHoliday & "�j"
	If ChkStr(oRS.Collect("WeeklyHolidayType")) <> "" Then
		If sHoliday <> "" Then sHoliday = sHoliday & "<br>"
		sHoliday = sHoliday & oRS.Collect("WeeklyHolidayType")
	End If
	If ChkStr(oRS.Collect("HolidayRemark")) <> "" Then
		If sHoliday <> "" Then sHoliday = sHoliday & "<br>"
		sHoliday = sHoliday & oRS.Collect("HolidayRemark")
	End If

	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">�]�Ή^�s��</th>"
	Response.Write "<td>" & sTransfer & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">�Ζ�����</th>"
	Response.Write "<td>"
	Response.Write sWorkPeriod

	If sWorkPeriodFlag = "3" And sWorkMonthPeriod <> "" Then
		'�Z��
		Response.Write " ( " & sWorkMonthPeriod & "���� )"
	End If

	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">�A�Ǝ���</th>"
	Response.Write "<td>" & sWorkTime & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">�c��</th>"
	Response.Write "<td>"
	Response.Write sOverWork

	If sOverWorkTimeMax <> "" Then
		Response.Write "(" & sOverWorkTimeMax & "�܂�)"
	End If

	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">�V�t�g(���)�ғ�</th>"
	Response.Write "<td>" & sWorkShift & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">�x��</th>"
	Response.Write "<td>" & sHoliday & "</td>"
	Response.Write "</tr>"
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̊�]�������������������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/12 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileHopeWelfare(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sStaffCode
	Dim sWelfareProgram

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	sWelfareProgram = ""
	If rRS.Collect("TrafficFeeFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@��ʔ�x��"
	If rRS.Collect("SocietyInsuranceFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@�Љ�ی�����"
	If rRS.Collect("SanatoriumFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@�ۗ{��"
	If rRS.Collect("EnterprisePensionFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@��ƔN��"
	If rRS.Collect("WealthShapeFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@���`���~"
	If rRS.Collect("StockOptionFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@�������x(�X�g�b�N�I�v�V����)"
	If rRS.Collect("RetirementPayFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@�ސE�����x"
	If rRS.Collect("ResidencePayFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@�Z��蓖"
	If rRS.Collect("FamilyPayFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@�Ƒ��蓖"
	If rRS.Collect("EmployeeDormitoryFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@�Ј���"
	If rRS.Collect("CompanyHouseFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@�Б�"
	If rRS.Collect("NewEmployeeTrainingFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@�V���Ј����C"
	If rRS.Collect("OverseasTrainingFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@�C�O���C"
	If rRS.Collect("OtherTrainingFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@�e�팤�C"
	If rRS.Collect("FlexTimeFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@�t���b�N�X�^�C��"

	If sWelfareProgram <> "" Then
		sWelfareProgram = Mid(sWelfareProgram, 2)

		Response.Write "<tr>"
		Response.Write "<th colspan=""2"">��������</th>"
		Response.Write "<td>" & sWelfareProgram & "</td>"
		Response.Write "</tr>"
	End If
End Function



'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̊�]�����������o��(smart)
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/12 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileHopeSmart(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vOrderCode)
	Dim sPattern

	If GetRSState(rRS) = False Then Exit Function

	If vUserType = "staff" Then
		sPattern = "pattern9"
	ElseIf vUserType = "company" Then
		sPattern = "pattern8"
	End If

	%>
	<table class="profileSmart smartBlock" style="display:none;">
        <thead>
            <tr>
                <th colspan="2">��]����</th>
            </tr>
        </thead>
        <tbody>
	
	<%

	'��]�Ζ��`��
	Call DspProfileHopeWorkingTypeSmart(rDB, rRS, vUserType, vUserID)
	'��]�Ǝ�
	Call DspProfileHopeIndustrySmart(rDB, rRS, vUserType, vUserID)
	'��]�E��
	Call DspProfileHopeJobTypeSmart(rDB, rRS, vUserType, vUserID)
	'��]�Ζ��n
	Call DspProfileHopeWorkingPlaceSmart(rDB, rRS, vUserType, vUserID)
	'���^����
	Call DspProfileHopeSalarySmart(rDB, rRS, vUserType, vUserID)
	'���ԁE����
	Call DspProfileHopeSpanSmart(rDB, rRS, vUserType, vUserID)
	'��������
	Call DspProfileHopeWelfareSmart(rDB, rRS, vUserType, vUserID)

	Response.Write "</tbody>"
	Response.Write "</table>" & vbCrLf
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̊�]�����Ζ��`�ԕ������o��(Smart)
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/12 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileHopeWorkingTypeSmart(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim sStaffCode
	Dim sWorkingType	: sWorkingType = ""

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	sSQL = "sp_GetDataWorkingType '" & sStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		sWorkingType = sWorkingType & oRS.Collect("WorkingTypeName")
		oRS.MoveNext
		If GetRSState(oRS) = True Then sWorkingType = sWorkingType & "�@"
	Loop
	Call RSClose(oRS)

	Response.Write "<tr>"
	Response.Write "<th>�Ζ��`��</th>"
	Response.Write "<td>"
	Response.Write sWorkingType
	Response.Write "</td>"
	Response.Write "</tr>"

	If G_USERTYPE = "dispatch" Then
		Response.Write "<tr>"
		Response.Write "<th colspan=""2"" class=""itNaiyo"">���Ј��Љ�\��h���̊�]</th>"
		response.write "</tr><tr>"
		Response.Write "<td colspan=""2"">"

		If rRS.Collect("TempToPermFlag") = "1" Then
			Response.Write "��]����"
		Else
			Response.Write "��]���Ȃ�"
		End If

		Response.Write "</td>"
		Response.Write "</tr>"
	End If
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̊�]�����Ǝ핔�����o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/12 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileHopeIndustrySmart(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sStaffCode

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	Response.Write "<tr>"
	Response.Write "<th>�Ǝ�</th>"
	Response.Write "<td>"

	sSQL = "sp_GetDataIndustryType '" & sStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		Response.Write oRS.Collect("IndustryTypeName")
		oRS.MoveNext
		If GetRSState(oRS) = True Then Response.Write "<br>"
	Loop
	Call RSClose(oRS)

	Response.Write "</td>"
	Response.Write "</tr>"
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̊�]�����Ǝ핔�����o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/12 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileHopeJobTypeSmart(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sStaffCode

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	Response.Write "<tr>"
	Response.Write "<th>�E��</th>"
	Response.Write "<td>"

	sSQL = "sp_GetDataHopeJobType '" & sStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		Response.Write oRS.Collect("JobTypeName")
		oRS.MoveNext
		If GetRSState(oRS) = True Then Response.Write "<br>"
	Loop
	Call RSClose(oRS)

	Response.Write "</td>"
	Response.Write "</tr>"
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̊�]�����Ζ��n�������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/12 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileHopeWorkingPlaceSmart(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sStaffCode

	Dim flgCommuteTime	: flgCommuteTime = False
	Dim flgStation		: flgStation = False
	Dim sArea			: sArea = ""
	Dim sPlace			: sPlace = ""
	Dim sHopeCommuteTime: sHopeCommuteTime = ""
	Dim sStation		: sStation = ""
	Dim sRailwayLine	: sRailwayLine = ""
	Dim iRow			: iRow = 1

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	sSQL = "sp_GetDataHopeWorkingPlace '" & sStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then iRow = iRow + 1
	Do While GetRSState(oRS) = True
		If InStr(sArea, oRS.Collect("Area")) = 0 Then
			'�d�����Ȃ��G���A���̂�
			sArea = sArea & "�@" & oRS.Collect("Area")
		End If
		sPlace = sPlace & oRS.Collect("WorkingPlace")
		oRS.MoveNext
		If GetRSState(oRS) = True Then sPlace = sPlace & "<br>"
	Loop
	Call RSClose(oRS)

	'��]�w
	sSQL = "sp_GetDataHopeCommuting '" & sStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		If sHopeCommuteTime <> "" Then sHopeCommuteTime = sHopeCommuteTime & "<br>"
		If ChkStr(oRS.Collect("HopeCommuteTime")) <> "" Then
			flgCommuteTime = True
			sHopeCommuteTime = oRS.Collect("HopeCommuteTime") & "��"
		End If

		If ChkStr(oRS.Collect("StationName")) <> "" Then
			flgStation = True
			If sStation <> "" Then sStation = sStation & "<br>"
			sStation = sStation & oRS.Collect("StationName") & "�w"

			If oRS.Collect("MinuteToStation") <> "" Then
				sStation = sStation & "(�w����" & oRS.Collect("MinuteToStation") & "���ȓ�)"
			End If
		End If
		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	If flgCommuteTime = True Then iRow = iRow + 1
	If flgStation = True Then iRow = iRow + 1

	'��]����
'	sSQL = "sp_GetDataHopeRailwayLine '" & sStaffCode & "'"
'	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
'	If GetRSState(oRS) = True Then iRow = iRow + 1
'	Do While GetRSState(oRS) = True
'		sRailwayLine = sRailwayLine & _
'			oRS.Collect("RailwayCompanyName") & _
'			"�@" & oRS.Collect("RailwayLineName") & "<br>"
'		oRS.MoveNext
'		If GetRSState(oRS) = True Then sRailwayLine = sRailwayLine & "<br>"
'	Loop
'	Call RSClose(oRS)

	Response.Write "<tr>"
	Response.Write "<th colspan=""2"" class=""itNaiyo"">�Ζ��n</th>"
	response.write "</tr><tr>"
	Response.Write "<th>��]��</th>"
	Response.Write "<td>" & sArea & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th>��]�Ζ��n</th>"
	Response.Write "<td>" & sPlace & "</td>"
	Response.Write "</tr>"

	If sHopeCommuteTime <> "" Then
		Response.Write "<tr>"
		Response.Write "<th>��]�ʋΎ���</th>"
		Response.Write "<td>" & sHopeCommuteTime & "</td>"
		Response.Write "</tr>"
	End If

	If sStation <> "" Then
		Response.Write "<tr>"
		Response.Write "<th>��]�w</th>"
		Response.Write "<td>" & sStation & "</td>"
		Response.Write "</tr>"
	End If
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̊�]�������^�������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/12 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileHopeSalarySmart(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sStaffCode
	Dim YMin
	Dim YMax
	Dim MMin
	Dim MMax
	Dim DMin
	Dim DMax
	Dim HMin
	Dim HMax
	Dim PercentagePay
	Dim Remark
	Dim iRow

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	If ChkStr(rRS.Collect("YearlyIncomeMin")) <> "" Then YMin = GetJapaneseYen(rRS.Collect("YearlyIncomeMin"))
	If ChkStr(rRS.Collect("YearlyIncomeMax")) <> "" Then YMax = GetJapaneseYen(rRS.Collect("YearlyIncomeMax"))
	If ChkStr(rRS.Collect("MonthlyIncomeMin")) <> "" Then MMin = GetJapaneseYen(rRS.Collect("MonthlyIncomeMin"))
	If ChkStr(rRS.Collect("MonthlyIncomeMax")) <> "" Then MMax = GetJapaneseYen(rRS.Collect("MonthlyIncomeMax"))
	If ChkStr(rRS.Collect("DailyIncomeMin")) <> "" Then DMin = GetJapaneseYen(rRS.Collect("DailyIncomeMin"))
	If ChkStr(rRS.Collect("DailyIncomeMax")) <> "" Then DMax = GetJapaneseYen(rRS.Collect("DailyIncomeMax"))
	If ChkStr(rRS.Collect("HourlyIncomeMin")) <> "" Then HMin = GetJapaneseYen(rRS.Collect("HourlyIncomeMin"))
	If ChkStr(rRS.Collect("HourlyIncomeMax")) <> "" Then HMax = GetJapaneseYen(rRS.Collect("HourlyIncomeMax"))
	PercentagePay = ChkStr(rRS.Collect("PercentagePayFlag"))
	Remark = ChkStr(rRS.Collect("IncomeRemark"))

	iRow = 0
	If YMin & YMax <> "" Then iRow = iRow + 1
	If MMin & MMax <> "" Then iRow = iRow + 1
	If DMin & DMax <> "" Then iRow = iRow + 1
	If HMin & HMax <> "" Then iRow = iRow + 1
	If PercentagePay <> "" Then iRow = iRow + 1
	If Remark <> "" Then iRow = iRow + 1

	If iRow > 0 Then
		Response.Write "<tr>"
		Response.Write "<th colspan=""2"" class=""itNaiyo"">��]���^</th>"
		response.write "</tr><tr>"
		If YMin & YMax <> "" Then
			Response.Write "<th>�N��</th>"
			Response.Write "<td>" & YMin & "�`" & YMax &"</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
		End If

		If MMin & MMax <> "" Then
			Response.Write "<th>����</th>"
			Response.Write "<td>" & MMin & "�`" & MMax & "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
		End If

		If DMin & DMax <> "" Then
			Response.Write "<th>����</th>"
			Response.Write "<td>" & DMin & "�`" & DMax & "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
		End If

		If HMin & HMax <> "" Then
			Response.Write "<th>����</th>"
			Response.Write "<td>" & HMin & "�`" & HMax & "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
		End If

		Response.Write "<th>����</th>"
		Response.Write "<td>"
		If PercentagePay = "1" Then
			Response.Write "��]����"
		ElseIf PercentagePay = "0" Then
			Response.Write "��]���Ȃ�"
		Else
			Response.Write "�������Ȃ�"
		End If
		Response.Write "</td>"
		Response.Write "</tr>"

		If Remark <> "" Then
			Response.Write "<tr>"
			Response.Write "<th>���l</th>"
			Response.Write "<td>" & Remark & "</td>"
			Response.Write "</tr>"
		End If
	Else
		'��]���^�����������ꍇ
		Response.Write "<tr>"
		Response.Write "<th colspan=""2"">��]���^</th>"
		response.write "</tr><tr>"
		Response.Write "<td colspan=""2""></td>"
		Response.Write "</tr>"
	End If
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̊�]�������ԁE���ԕ������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/12 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileHopeSpanSmart(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sStaffCode
	Dim sTransfer
	Dim sWorkPeriod
	Dim sWorkPeriodFlag
	Dim sWorkMonthPeriod
	Dim sWorkTime
	Dim sWSTime
	Dim sWETime
	Dim sOverWork
	Dim sOverWorkTimeMax
	Dim sWorkShift
	Dim sHoliday

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	sTransfer = ChkStr(rRS.Collect("TransferFlag"))
	Select Case sTransfer
		Case "1":	sTransfer = "��"
		Case "0":	sTransfer = "�s��"
		Case Else:	sTransfer = "�������Ȃ�"
	End Select

	sWorkPeriod = ChkStr(oRS.Collect("WorkPeriod"))
	sWorkPeriodFlag = ChkStr(oRS.Collect("WorkPeriodFlag"))
	sWorkMonthPeriod = ChkStr(oRS.Collect("WorkMonthPeriod"))
	sWSTime = ChkStr(oRS.Collect("WorkStartTime"))
	sWETime = ChkStr(oRS.Collect("WorkEndTime"))
	If sWSTime & sWETime <> "" Then
		If sWSTime <> "" Then sWorkTime = sWorkTime & Left(sWSTime, 2) & ":" & Right(sWSTime, 2)
		sWorkTime = sWorkTime & "�`"
		If sWETime <> "" Then sWorkTime = sWorkTime & Left(sWETime, 2) & ":" & Right(sWETime, 2)
	End If

	sOverWork = ChkStr(rRS.Collect("OverWorkFlag"))
	Select Case sOverWork
		Case "1":	sOverWork = "��"
		Case "0":	sOverWork = "�s��"
		Case Else:	sOverWork = "�������Ȃ�"
	End Select
	sOverWorkTimeMax = ChkStr(rRS.Collect("OverWorkTimeMax"))
	If sOverWorkTimeMax <> "" Then sOverWorkTimeMax = Left(sOverWorkTimeMax, 2) & ":" & Right(sOverWorkTimeMax, 2)

	sWorkShift = ChkStr(rRS.Collect("WorkShiftFlag"))
	Select Case sWorkShift
		Case "1":	sWorkShift = "��"
		Case "0":	sWorkShift = "�s��"
		Case Else:	sWorkShift = "�������Ȃ�"
	End Select

	sHoliday = ""
	If oRS.Collect("MonHolidayFlag") = "1" Then sHoliday = sHoliday & "��"
	If oRS.Collect("TueHolidayFlag") = "1" Then sHoliday = sHoliday & "��"
	If oRS.Collect("WedHolidayFlag") = "1" Then sHoliday = sHoliday & "��"
	If oRS.Collect("ThuHolidayFlag") = "1" Then sHoliday = sHoliday & "��"
	If oRS.Collect("FriHolidayFlag") = "1" Then sHoliday = sHoliday & "��"
	If oRS.Collect("SatHolidayFlag") = "1" Then sHoliday = sHoliday & "�y"
	If oRS.Collect("SunHolidayFlag") = "1" Then sHoliday = sHoliday & "��"
	If oRS.Collect("PublicHolidayFlag") = "1" Then sHoliday = sHoliday & "�j"
	If ChkStr(oRS.Collect("WeeklyHolidayType")) <> "" Then
		If sHoliday <> "" Then sHoliday = sHoliday & "<br>"
		sHoliday = sHoliday & oRS.Collect("WeeklyHolidayType")
	End If
	If ChkStr(oRS.Collect("HolidayRemark")) <> "" Then
		If sHoliday <> "" Then sHoliday = sHoliday & "<br>"
		sHoliday = sHoliday & oRS.Collect("HolidayRemark")
	End If

	Response.Write "<tr>"
	Response.Write "<th>�]�Ή^�s��</th>"
	Response.Write "<td>" & sTransfer & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th>�Ζ�����</th>"
	Response.Write "<td>"
	Response.Write sWorkPeriod

	If sWorkPeriodFlag = "3" And sWorkMonthPeriod <> "" Then
		'�Z��
		Response.Write " ( " & sWorkMonthPeriod & "���� )"
	End If

	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th>�A�Ǝ���</th>"
	Response.Write "<td>" & sWorkTime & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th>�c��</th>"
	Response.Write "<td>"
	Response.Write sOverWork

	If sOverWorkTimeMax <> "" Then
		Response.Write "(" & sOverWorkTimeMax & "�܂�)"
	End If

	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th>�V�t�g(���)�ғ�</th>"
	Response.Write "<td>" & sWorkShift & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th>�x��</th>"
	Response.Write "<td>" & sHoliday & "</td>"
	Response.Write "</tr>"
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̊�]�������������������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/12 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileHopeWelfareSmart(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sStaffCode
	Dim sWelfareProgram

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	sWelfareProgram = ""
	If rRS.Collect("TrafficFeeFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@��ʔ�x��"
	If rRS.Collect("SocietyInsuranceFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@�Љ�ی�����"
	If rRS.Collect("SanatoriumFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@�ۗ{��"
	If rRS.Collect("EnterprisePensionFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@��ƔN��"
	If rRS.Collect("WealthShapeFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@���`���~"
	If rRS.Collect("StockOptionFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@�������x(�X�g�b�N�I�v�V����)"
	If rRS.Collect("RetirementPayFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@�ސE�����x"
	If rRS.Collect("ResidencePayFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@�Z��蓖"
	If rRS.Collect("FamilyPayFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@�Ƒ��蓖"
	If rRS.Collect("EmployeeDormitoryFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@�Ј���"
	If rRS.Collect("CompanyHouseFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@�Б�"
	If rRS.Collect("NewEmployeeTrainingFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@�V���Ј����C"
	If rRS.Collect("OverseasTrainingFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@�C�O���C"
	If rRS.Collect("OtherTrainingFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@�e�팤�C"
	If rRS.Collect("FlexTimeFlag") = "1" Then sWelfareProgram = sWelfareProgram & "�@�t���b�N�X�^�C��"

	If sWelfareProgram <> "" Then
		sWelfareProgram = Mid(sWelfareProgram, 2)

		Response.Write "<tr>"
		Response.Write "<th>��������</th>"
		Response.Write "<td>" & sWelfareProgram & "</td>"
		Response.Write "</tr>"
	End If
End Function




'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̓]�E�̏������i�L�����A�I�����A�i���C�U�[���j���o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�F
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2009/09/24 LIS T.Ezaki �쐬
'******************************************************************************
Function DspCareerAnalyzer(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sStaffCode
	Dim sPattern

	If vUserType = "staff" Then
		sPattern = "pattern9"
	ElseIf vUserType = "company" Then
		sPattern = "pattern8"
	End If
	
	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")
	
	sSQL = "select "
	sSQL = sSQL & " IdealIndustryText"
	sSQL = sSQL & ",IdealIndustryPriority"
	sSQL = sSQL & ",IdealIndustryDistance"
	sSQL = sSQL & ",IdealPositionText"
	sSQL = sSQL & ",IdealPositionPriority"
	sSQL = sSQL & ",IdealPositionDistance"
	sSQL = sSQL & ",IdealJobText"
	sSQL = sSQL & ",IdealJobPriority"
	sSQL = sSQL & ",IdealJobDistance"
	sSQL = sSQL & ",IdealCustomText"
	sSQL = sSQL & ",IdealCustomPriority"
	sSQL = sSQL & ",IdealCustomDistance"
	sSQL = sSQL & ",IdealServiceText"
	sSQL = sSQL & ",IdealServicePriority"
	sSQL = sSQL & ",IdealServiceDistance"
	sSQL = sSQL & ",IdealRelationsText"
	sSQL = sSQL & ",IdealRelationsPriority"
	sSQL = sSQL & ",IdealRelationsDistance"
	sSQL = sSQL & ",IdealFutureText"
	sSQL = sSQL & ",IdealFuturePriority"
	sSQL = sSQL & ",IdealFutureDistance"
	sSQL = sSQL & ",IdealWorkareaText"
	sSQL = sSQL & ",IdealWorkareaPriority"
	sSQL = sSQL & ",IdealWorkareaDistance"
	sSQL = sSQL & ",IdealTrainingText"
	sSQL = sSQL & ",IdealTrainingPriority"
	sSQL = sSQL & ",IdealTrainingDistance"
	sSQL = sSQL & ",IsNull(CareerGoolType,'') as CareerGoolType"
	sSQL = sSQL & ",IsNull(CareerGoolEtc,'') as CareerGoolEtc"
	sSQL = sSQL & ",IsNull(CareerGoolDetail,'') as CareerGoolDetail"
	sSQL = sSQL & ",Publicflag"
	sSQL = sSQL & " from P_CareerAnalyzer"
	sSQL = sSQL & " Where StaffCode = '" & sStaffCode & "' and Publicflag = 1"
	
	flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
	If GetRSState(oRS) = true then
		response.write "<table class=""" & sPattern & " cw"" style=""margin-bottom:15px;"">"
		response.write "<thead><tr><th colspan=""4"" style=""width:588px;"">�����ȗ��z��</th></tr></thead>"
		response.write "<tbody>"
		response.write "<tr>"
		response.write "<th style=""width:188px;"">���</th>"
		response.write "<th style=""text-align:left"">���z��</th>"
		response.write "<th style=""text-align:center"">�D��x</th>"
		response.write "<th style=""text-align:center"">������</th>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<th>�ƊE�E�Ǝ�̗��z��</th>"
		response.write "<td>" & oRS.Collect("IdealIndustryText") & "</td>"
		response.write "<td style=""text-align:center"">" & oRS.Collect("IdealIndustryPriority") & "</td>"
		response.write "<td style=""text-align:center"">"
		select case oRS.Collect("IdealIndustryDistance")
			case "1"	:	response.write "����"
			case "2"	:	response.write "��≓��"
			case "3"	:	response.write "�ӂ�"
			case "4"	:	response.write "���߂�"
			case "5"	:	response.write "�߂�" 
		end select
		response.write "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<th>�|�W�V�����̗��z��</th>"
		response.write "<td>" & oRS.Collect("IdealPositionText") & "</td>"
		response.write "<td style=""text-align:center"">" & oRS.Collect("IdealPositionPriority") & "</td>"
		response.write "<td style=""text-align:center"">"
		select case oRS.Collect("IdealPositionDistance")
			case "1"	:	response.write "����"
			case "2"	:	response.write "��≓��"
			case "3"	:	response.write "�ӂ�"
			case "4"	:	response.write "���߂�"
			case "5"	:	response.write "�߂�" 
		end select
		response.write "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<th>�E��E�d�����e�̗��z��</th>"
		response.write "<td>" & oRS.Collect("IdealJobText") & "</td>"
		response.write "<td style=""text-align:center"">" & oRS.Collect("IdealJobPriority") & "</td>"
		response.write "<td style=""text-align:center"">"
		select case oRS.Collect("IdealJobDistance")
			case "1"	:	response.write "����"
			case "2"	:	response.write "��≓��"
			case "3"	:	response.write "�ӂ�"
			case "4"	:	response.write "���߂�"
			case "5"	:	response.write "�߂�" 
		end select
		response.write "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<th>�Е��̗��z��</th>"
		response.write "<td>" & oRS.Collect("IdealCustomText") & "</td>"
		response.write "<td style=""text-align:center"">" & oRS.Collect("IdealCustomPriority") & "</td>"
		response.write "<td style=""text-align:center"">"
		select case oRS.Collect("IdealCustomDistance")
			case "1"	:	response.write "����"
			case "2"	:	response.write "��≓��"
			case "3"	:	response.write "�ӂ�"
			case "4"	:	response.write "���߂�"
			case "5"	:	response.write "�߂�" 
		end select
		response.write "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<th>���^�E�ҋ��̗��z��</th>"
		response.write "<td>" & oRS.Collect("IdealServiceText") & "</td>"
		response.write "<td style=""text-align:center"">" & oRS.Collect("IdealServicePriority") & "</td>"
		response.write "<td style=""text-align:center"">"
		select case oRS.Collect("IdealServiceDistance")
			case "1"	:	response.write "����"
			case "2"	:	response.write "��≓��"
			case "3"	:	response.write "�ӂ�"
			case "4"	:	response.write "���߂�"
			case "5"	:	response.write "�߂�" 
		end select
		response.write "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<th>�l�Ԋ֌W�̗��z��</th>"
		response.write "<td>" & oRS.Collect("IdealRelationsText") & "</td>"
		response.write "<td style=""text-align:center"">" & oRS.Collect("IdealRelationsPriority") & "</td>"
		response.write "<td style=""text-align:center"">"
		select case oRS.Collect("IdealRelationsDistance")
			case "1"	:	response.write "����"
			case "2"	:	response.write "��≓��"
			case "3"	:	response.write "�ӂ�"
			case "4"	:	response.write "���߂�"
			case "5"	:	response.write "�߂�" 
		end select
		response.write "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<th>��Џ������̗��z��</th>"
		response.write "<td>" & oRS.Collect("IdealFutureText") & "</td>"
		response.write "<td style=""text-align:center"">" & oRS.Collect("IdealFuturePriority") & "</td>"
		response.write "<td style=""text-align:center"">"
		select case oRS.Collect("IdealFutureDistance")
			case "1"	:	response.write "����"
			case "2"	:	response.write "��≓��"
			case "3"	:	response.write "�ӂ�"
			case "4"	:	response.write "���߂�"
			case "5"	:	response.write "�߂�" 
		end select
		response.write "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<th>�Ζ��n�̗��z��</th>"
		response.write "<td>" & oRS.Collect("IdealWorkareaText") & "</td>"
		response.write "<td style=""text-align:center"">" & oRS.Collect("IdealWorkareaPriority") & "</td>"
		response.write "<td style=""text-align:center"">"
		select case oRS.Collect("IdealWorkareaDistance")
			case "1"	:	response.write "����"
			case "2"	:	response.write "��≓��"
			case "3"	:	response.write "�ӂ�"
			case "4"	:	response.write "���߂�"
			case "5"	:	response.write "�߂�" 
		end select
		response.write "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<th>���猤�C�̗��z��</th>"
		response.write "<td>" & oRS.Collect("IdealTrainingText") & "</td>"
		response.write "<td style=""text-align:center"">" & oRS.Collect("IdealTrainingPriority") & "</td>"
		response.write "<td style=""text-align:center"">"
		select case oRS.Collect("IdealTrainingDistance")
			case "1"	:	response.write "����"
			case "2"	:	response.write "��≓��"
			case "3"	:	response.write "�ӂ�"
			case "4"	:	response.write "���߂�"
			case "5"	:	response.write "�߂�" 
		end select
		response.write "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<th>�L�����A�̃S�[���C���[�W"
		response.write "</th>"
		response.write "<td colspan=""3"">"
		response.write "<strong style=""font-size:16px;"">"
		Select Case oRS.Collect("CareerGoolType")
			Case "1"	:	Response.Write "�X�L���T���u��"
			Case "2"	:	Response.Write "�}�l�[�W�����g�u��"
			Case "3"	:	Response.Write "�Ɨ��u��"
			Case "4"	:	Response.Write oRS.Collect("CareerGoolEtc")
		End Select
		response.write "</strong><br>"
		Response.Write Replace(oRS.Collect("CareerGoolDetail"),vbCrLf,"<br>")
		response.write "</td>"
		response.write "</tr>"
		response.write "</tbody>"
		response.write "</table>"
		Response.Write "<p style=""text-align:right;"" class=""m0""><a class=""stext"" href=""#pagetop"">���y�[�WTOP��</a></p>" & vbCrLf
	end if
	Call RSClose(oRS)
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̍ŐV���M���[���󋵕������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/12 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileMail(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	'DB
	Dim dbStaffCode
	Dim dbSendDay1
	Dim dbSubject1
	Dim dbBody1
	Dim dbSendDay2
	Dim dbSubject2
	Dim dbBody2

	If GetRSState(rRS) = False Then Exit Function

	dbStaffCode = rRS.Collect("StaffCode")

	If vUserType = "company" Or vUserType = "dispatch" Then
		sSQL = "EXEC up_DtlMailHistory_Staff '" & vUserID & "', '" & dbStaffCode & "', '" & vOrderCode & "';"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			dbSendDay1 = oRS.Collect("SendDay1")
			dbSubject1 = oRS.Collect("Subject1")
			dbBody1 = ChkStr(oRS.Collect("Body1"))
			dbSendDay2 = oRS.Collect("SendDay2")
			dbSubject2 = oRS.Collect("Subject2")
			dbBody2 = ChkStr(oRS.Collect("Body2"))

			Response.Write "<table class=""pattern2 cw"" border=""0"" style=""margin-bottom:15px;"">"
			Response.Write "<colgroup>"
			Response.Write "<col style=""width:300px;"">"
			Response.Write "<col style=""width:300px;"">"
			Response.Write "<colgroup>"
			Response.Write "<thead>"
			Response.Write "<tr>"
			Response.Write "<th colspan=""2"" style=""text-align:center;"">���[���ŐV��</th>"
			Response.Write "</tr>"
			Response.Write "</thead>"
			Response.Write "<tbody>"
			Response.Write "<tr>"
			Response.Write "<th style=""text-align:center;"">�M�Ђ����M��������</th>"
			Response.Write "<th style=""text-align:center;"">���E�҂����M��������</th>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td style=""vertical-align:top;"">"
			If ChkStr(dbSendDay1) <> "" Then
				Response.Write "�y���M�����z<br>" & dbSendDay1
				Response.Write "<div class=""line1""></div>"
				Response.Write "�y�����z<br>" & dbSubject1
				Response.Write "<div class=""line1""></div>"
				Response.Write "�y�{���z<br>" & Replace(dbBody1,vbCrLf,"<br>")
			End If
			Response.Write "</td>"
			Response.Write "<td style=""vertical-align:top;"">"
			If ChkStr(dbSendDay2) <> "" Then
				Response.Write "�y���M�����z<br>" & dbSendDay2
				Response.Write "<div class=""line1""></div>"
				Response.Write "�y�����z<br>" & dbSubject2
				Response.Write "<div class=""line1""></div>"
				Response.Write "�y�{���z<br>" & Replace(dbBody2,vbCrLf,"<br>")
			End If
			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "</tbody>"
			Response.Write "</table>"
			Response.Write "<p style=""text-align:right;"" class=""m0""><a class=""stext"" href=""#pagetop"">���y�[�WTOP��</a></p>" & vbCrLf
		End If
		Call RSClose(oRS)
	End If
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̋��E�҃R�[�h���o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/12 LIS K.Kokubo �쐬
'�@�@�@�F2008/09/09 LIS K.Kokubo ���E�Җ����͑Ή�
'******************************************************************************
Function DspProfileStaffCode(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim oRS
	Dim sSQL
	Dim flgQE
	Dim sError

	Dim dbStaffName
	Dim sStaffCode
	Dim sTableClass

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	sSQL = "EXEC up_DtlCMPStaffName '" & vUserID & "', '" & sStaffCode & "';"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		dbStaffName = ChkStr(oRS.Collect("StaffName"))
	End If
	Call RSClose(oRS)

	sTableClass = "pattern9"
	If vUserType = "company" Or vUserType = "dispatch" Then sTableClass = "pattern8"

	Response.Write "<div class=""center"">"
	If dbStaffName <> "" Then
		Response.Write dbStaffName & "&nbsp;(���E�҃R�[�h�F" & sStaffCode & ")"
	Else
		Response.Write sStaffCode & "(���ID)"
	End If

	Response.Write "</div>" & vbCrLf
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̊e�ҏW�{�^�����o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/16 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileEditButton(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Response.Write "<form id=""frmEdit"" action="""" method=""post"">"
	Response.Write "<input type=""hidden"" name=""CONF_StaffCode"" value=""" & vUserID & """>"
	Response.Write "<h2>�e��ҏW�{�^��</h2>"
	Response.Write "���ҏW����ꍇ�͈ȉ��̃{�^���������ĉ������B<br>"
	Response.Write "<table class=""cw profileEdit"" border=""1"" bordercolor=""#ff9999"" style=""font-size:10px; width:745px;margin-left:-5px;"">"
	Response.Write "<tr>"
	Response.Write "<td align=""center"" valign=""top"">"
	Response.Write "�������E�E���o����(��ʁj�E�X�J�E�g���<br>"
	Response.Write "<input type=""image"" src=""/img/staff/person_detail_edit1.jpg"" id=""button1"" value=""�l�f�[�^"" onclick=""document.forms.frmEdit.action = './person_edit1.asp'; document.forms.frmEdit.submit();"">"
	Response.Write "<input type=""image"" src=""/img/staff/person_detail_edit2.jpg"" id=""button2"" value=""�w��E��"" onclick=""document.forms.frmEdit.action = './person_edit2.asp'; document.forms.frmEdit.submit();"" style=""padding-left:3px;"">"
	Response.Write "<input type=""image"" src=""/img/staff/person_detail_edit3.jpg"" id=""button3"" value=""���i�E�X�L��"" onclick=""document.forms.frmEdit.action = './person_edit3.asp'; document.forms.frmEdit.submit();"" style=""padding-left:3px;"">"
	Response.Write "<input type=""image"" src=""/img/staff/person_detail_edit4.jpg"" id=""button4"" value=""IT�n&#10;&#13;�J���X�L��"" onclick=""document.forms.frmEdit.action = './person_edit4.asp'; document.forms.frmEdit.submit();"" style=""padding-left:3px;"">"
	Response.Write "<input type=""image"" src=""/img/staff/person_detail_edit9.jpg"" id=""button9"" value=""��w�X�L��"" onclick=""document.forms.frmEdit.action = './person_edit9.asp'; document.forms.frmEdit.submit();"" style=""padding-left:3px;"">"
	Response.Write "<input type=""image"" src=""/img/staff/person_detail_edit7.jpg"" id=""button7"" value=""���Ȃo�q&#10;&#13;�u�]���@��"" onclick=""document.forms.frmEdit.action = './person_edit7.asp'; document.forms.frmEdit.submit();"" style=""padding-left:3px;"">"
	Response.Write "<input type=""image"" src=""/img/staff/person_detail_edit8.jpg"" id=""button8"" value=""���ӕ��쓙"" onclick=""document.forms.frmEdit.action = './person_edit8.asp'; document.forms.frmEdit.submit();"" style=""padding-left:3px;"">"
	Response.Write "</td>"
	Response.Write "<td align=""center"" valign=""top"">"
	Response.Write "�E���o����(IT)<br>"
	Response.Write "<input type=""image"" src=""/img/staff/person_detail_edit5.jpg"" id=""button5"" value=""IT�n&#10;&#13;�E���ڍ�"" onclick=""document.forms.frmEdit.action = './person_edit5.asp'; document.forms.frmEdit.submit();"">"
	Response.Write "</td>"
	Response.Write "<td align=""center"" valign=""top"">"
	Response.Write "�X�J�E�g���<br>"
    '2015/09/01 �����N�ύX
	'Response.Write "<input type=""image"" src=""/img/staff/person_detail_edit6.jpg"" id=""button6"" value=""�{�l��]"" onclick=""document.forms.frmEdit.action = './person_edit6.asp'; document.forms.frmEdit.submit();"">"
    Response.Write "<input type=""image"" src=""/img/staff/person_detail_edit6.jpg"" id=""button6"" value=""�{�l��]"" onclick=""document.forms.frmEdit.action = './step2a.asp'; document.forms.frmEdit.submit();"">"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "</form>" & vbCrLf

	Response.Write "<p>"
	Response.Write "���������𑽂����͂��đ��l�ƍ������܂��傤�I<br>"
	Response.Write "<span style=""color:red;font-weight:bold;font-size:14px;line-height:font-weight:14px;"">�����L���e�͊�Ƒ����X�J�E�g����ۂɕ\���������e�Ɠ���ł��B</span><br>"
	Response.Write "<span style=""color:#ff0000;"">�����͂�����e����͂���΂��̗p����₷���Ȃ�܂��B�Ȃ��l�����ł����񂪏o�Ă��Ȃ������m�F���������B</span>"
	Response.Write "</p>" & vbCrLf
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̍ŏI�X�V�����o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/16 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileUpdateDay(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	If GetRSState(rRS) = False Then Exit Function

	Response.Write "<table class=""cw"" border=""0"">"
	Response.Write "<tr>"
	Response.Write "<td style=""text-align:right;"">�ŏI�X�V���F" & GetDateStr(rRS.Collect("UpdateDay"), "/") & "</td>"
	Response.Write "</tr>"
	Response.Write "</table>" & vbCrLf
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̐E����͖����ɑ΂��镶�����o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/16 LIS K.Kokubo �쐬
'******************************************************************************
Function DspProfileAttentionCareerHistory(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	If GetRSState(rRS) = False Then Exit Function
%>
<div style="margin-top:0px;">
	<table class="cw" border="0" style="width:270px; float:left; margin-bottom:10px;">
		<thead>
		<tr>
			<th style="padding:5px;border:1px solid #ff0000; border-width:1px 1px 0px 1px;background-color:#ffdddd;">�E������͂��܂��傤�I</th>
		</tr>
		</thead>
		<tbody>
		<tr>
			<td style="padding:5px;border:1px solid #ff0000;background-color:#ffeeee;">
				<p>
					�E���͋��l��Ƃ��ƂĂ����ڂ��Ă��鍀�ڂł��B<br>
					�E���̓��͂��������́A<br>
					<b>(1)�X�J�E�g����ɂ���</b><br>
					<b>(2)���債�Ă����ޑI�l�ŗ����Ă��܂�</b><br>
					���̉\���������Ȃ�܂��B<br>
				</p>
			</td>
		</tr>
		</tbody>
	</table>

	<table class="cw" border="0" style="width:320px; float:right;">
		<thead>
		<tr>
			<th style="padding:5px;border:1px solid #ff0000; border-width:1px 1px 0px 1px;background-color:#ffdddd;">�o�^�����ɂ��A�ȉ��@�\�������p�\�ł�</th>
		</tr>
		</thead>
		<tbody>
		<tr>
			<td style="padding:5px;border:1px solid #ff0000;background-color:#ffeeee;">
				<p>
					��<a href="/order/order_search_detail.asp" title="���d���ڍ׌���">�D���Ȏd���ɉ���</a>�ł���B<br>
					��<a href="/staff/s_resume.asp" title="������">������</a>�E<a href="/staff/s_careersheet.asp" title="�E���o����">�E���o����</a>�̈�����ł���B<br>
					��<a href="/s_contents/motive_index.asp" title="�u�]���@">�u�]���@</a>�E<a href="/s_contents/s_jikopr.asp" title="����PR">����PR</a>�̍쐬�x���c�[�����g�p�ł���B<br>
					���S���w�Ɋ�Â���<a href="/s_contents/s_mynavi.asp" title="�K�E�f�f">�K�E�f�f</a>���󂯂���B<br>
					��<a href="/s_contents/enquete.asp" title="�������ƃA���P�[�g">�������ƃA���P�[�g</a>�ɎQ���ł���B<br>
				</p>
			</td>
		</tr>
		</tbody>
	</table>
    <p style="text-align:right;" class="m0"><a class="stext" href="#pagetop">���y�[�WTOP��</a></p>
    <br clear="all">
</div>
<%
End Function

'******************************************************************************
'�T�@�v�F���[���}�K�W������̃v���t�B�[���y�[�W�ւ̃A�N�Z�X�����O�ɋL�^
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2007/04/16 LIS K.Kokubo �쐬
'******************************************************************************
Function RegMailMagazineAccess(ByRef rDB, ByVal vUserID, ByVal vMailMagazineID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sSuspensionFlag

	RegMailMagazineAccess = False
	sSuspensionFlag = "0"
	If vMailMagazineID <> "" Then
		RegMailMagazineAccess = True
		sSQL = "up_MailMagazineAccess '" & vMailMagazineID & "', '" & vUserID & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			If oRS.Collect("SuspensionFlag") = "1" Then RegMailMagazineAccess = False
		Else
			RegMailMagazineAccess = False
		End If
		Call RSClose(oRS)
	End If
End Function

'******************************************************************************
'�T�@�v�F���X�V�̒ʒm�����X�̎Ј��Ƀ��[������
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�@�@�@�FvMailServer	�F���[���T�[�o�[
'�@�@�@�FvFrom			�F���M��
'�@�@�@�FvSubject		�F����
'�@�@�@�FvBody			�F���e
'�g�p���F
'���@�l�F
'���@���F2007/04/16 LIS K.Kokubo �쐬
'�@�@�@�F2010/10/22 LIS K.Kokubo �k�h�r�����Ј��ɂ̂݃��[���𑗂�悤�ɏC��
'******************************************************************************
Function SendMailStaffEdit(ByRef rDB, ByVal vUserID, ByVal vMailServer, ByVal vFrom, ByVal vSubject, ByVal vBody)
	Dim sSQL,oRS,flgQE,sSQLErr
	Dim sRet,sBody,sLisEmployeeMail

	sLisEmployeeMail = ""
	'<�Ώۋ��E�҂��E�H�b�`���X�g�ɓo�^���Ă��郊�X�����Ј��̃��[���A�h���X�ꗗ���擾>
	sSQL = "EXEC up_LstLISWatchList '" & G_USERID & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sSQLErr)
	Do While GetRSState(oRS) = True
		If sLisEmployeeMail <> "" Then sLisEmployeeMail = sLisEmployeeMail & vbTab
		sLisEmployeeMail = sLisEmployeeMail & ChkStr(oRS.Collect("MailAddress"))
		oRS.MoveNext
	Loop
	'</�Ώۋ��E�҂��E�H�b�`���X�g�ɓo�^���Ă��郊�X�����Ј��̃��[���A�h���X�ꗗ���擾>

	sBody = ""
	sBody = sBody & "�����E�҃R�[�h�@[" & G_USERID & "]" & vbCrLf
	sBody = sBody & HTTP_BI_CURRENTURL & "staff/Staff_detail.asp?staffcode=" & G_USERID & vbCrLf & vbCrLf 
	sBody = sBody & vBody

	sRet = SndMail(vMailServer, sLisEmployeeMail, vFrom, vSubject, sBody, "")
End Function

'******************************************************************************
'�T�@�v�F��Ƃ����E�ҏڍׂ��{�������烍�O�ɏ�������
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FvCompanyCode	�F���O�C������ƃR�[�h
'�@�@�@�FvOrderCode		�F���O�C������Ƃ����E�҂��������Ă��邨�d���̏��R�[�h
'�@�@�@�FvStaffCode		�F���E�҃R�[�h
'�g�p���F
'���@�l�F
'�X�@�V�F2007/10/22 LIS K.Kokubo �쐬
'******************************************************************************
Function RegAccessHistoryStaff(ByRef rDB, ByVal vCompanyCode, ByVal vOrderCode, ByVal vStaffCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	RegAccessHistoryStaff = False

	If vCompanyCode <> "" And vOrderCode <> "" And vStaffCode <> "" And vOrderCode <> vStaffCode Then
		sSQL = "up_RegLOG_AccessHistoryStaff '" & vOrderCode & "', '" & vStaffCode & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

		RegAccessHistoryStaff = flgQE
	End If
End Function

'******************************************************************************
'�T�@�v�F��Ƃ����E�ҏڍׂ��{�������烍�O�ɏ�������
'���@���FrDB				�F�ڑ�����DBConnection
'�@�@�@�FvCompanyCode		�F���O�C������ƃR�[�h
'�@�@�@�FvOrderCode			�F���O�C������Ƃ����E�҂��������Ă��邨�d���̏��R�[�h
'�@�@�@�FvStaffCodeArray	�F���E�҃R�[�h�̔z��
'�@�@�@�FvPage				�F�\�����̈ꗗ�̃y�[�W��
'�g�p���F
'���@�l�F
'�X�@�V�F2007/11/12 LIS K.Kokubo �쐬
'******************************************************************************
Function RegAccessHistoryStaffList(ByRef rDB, ByVal vCompanyCode, ByVal vOrderCode, ByVal vStaffCodeArray)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sWhere

	Dim flgResult
	Dim idx
	Dim sDeclare
	Dim sParams

	flgResult = False

	'�Ώۋ��E�҂̗L�����`�F�b�N����, INSERT��, UPDATE�� �̍쐬
	If vCompanyCode <> "" And vOrderCode <> "" And UBound(vStaffCodeArray) >= 0 Then
		'**************************************************************************
		'�@UPDATE���쐬 start
		'--------------------------------------------------------------------------
		sWhere = ""
		sDeclare = "@vOrderCode CHAR(8),@vCompanyCode CHAR(8)"
		sParams = ",@vOrderCode = N'" & vOrderCode & "',@vCompanyCode = N'" & vCompanyCode & "'"

		For idx = LBound(vStaffCodeArray) To UBound(vStaffCodeArray)
			sDeclare = sDeclare & ",@vStaffCode" & idx & " CHAR(8)"
			sParams = sParams & ",@vStaffCode" & idx & " = N'" & vStaffCodeArray(idx) & "'"

			If sWhere <> "" Then sWhere = sWhere & ","
			sWhere = sWhere & "@vStaffCode" & idx
		Next

		sSQL = "" & _
			"UPDATE LOG_AccessHistoryStaffList " & _
			"SET YYYYMM = CONVERT(VARCHAR(4), YEAR(GETDATE())) + RIGHT('0' + CONVERT(VARCHAR(2), MONTH(GETDATE())), 2) " & _
				",UpdateDay = GETDATE() " & _
			"WHERE YYYYMM = CONVERT(VARCHAR(4), YEAR(GETDATE())) + RIGHT('0' + CONVERT(VARCHAR(2), MONTH(GETDATE())), 2) " & _
				"AND OrderCode = @vOrderCode " & _
				"AND StaffCode IN (" & sWhere & ") "

		sSQL = Replace(sSQL, "'", "''")
		sSQL = "sp_executesql N'" & sSQL & "', N'" & sDeclare & "'" & sParams
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		flgResult = flgQE
		'--------------------------------------------------------------------------
		'�@UPDATE���쐬 end
		'**************************************************************************

		'**************************************************************************
		'�AINSERT���쐬 start
		'--------------------------------------------------------------------------
		sSQL = ""
		sWhere = ""
		sDeclare = "@vOrderCode CHAR(8),@vCompanyCode CHAR(8)"
		sParams = ",@vOrderCode = N'" & vOrderCode & "',@vCompanyCode = N'" & vCompanyCode & "'"

		For idx = LBound(vStaffCodeArray) To UBound(vStaffCodeArray)
			sDeclare = sDeclare & ",@vStaffCode" & idx & " CHAR(8)"
			sParams = sParams & ",@vStaffCode" & idx & " = N'" & vStaffCodeArray(idx) & "'"

			If sSQL <> "" Then sSQL = sSQL & "UNION "
			sSQL = sSQL & _
				"SELECT CONVERT(VARCHAR(4), YEAR(GETDATE())) + RIGHT('0' + CONVERT(VARCHAR(2), MONTH(GETDATE())), 2) AS YYYYMM " & _
					",@vStaffCode" & idx & " AS StaffCode " & _
					",@vOrderCode AS OrderCode " & _
					",@vCompanyCode AS CompanyCode " & _
					",GETDATE() AS RegistDay "

			If sWhere <> "" Then sWhere = sWhere & ","
			sWhere = sWhere & "@vStaffCode" & idx
		Next

		sSQL = "INSERT INTO LOG_AccessHistoryStaffList " & _
			"SELECT INS.YYYYMM, INS.StaffCode, INS.OrderCode, INS.CompanyCode, INS.RegistDay, INS.RegistDay " & _
			"FROM (" & sSQL & ") AS INS " & _
			"WHERE NOT EXISTS( " & _
					"SELECT * " & _
					"FROM LOG_AccessHistoryStaffList AS NEX " & _
					"WHERE INS.YYYYMM = NEX.YYYYMM " & _
						"AND INS.OrderCode = NEX.OrderCode " & _
						"AND NEX.StaffCode IN (" & sWhere & ") " & _
				")"

		sSQL = Replace(sSQL, "'", "''")
		sSQL = "sp_executesql N'" & sSQL & "', N'" & sDeclare & "'" & sParams
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		flgResult = flgQE
		'--------------------------------------------------------------------------
		'�AINSERT���쐬 end
		'**************************************************************************
	End If

	RegAccessHistoryStaffList = flgResult
End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̊�]�A�����ԑѕ������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2015/07/17 LIS K.Kimura �쐬
'******************************************************************************
Function DspProfileHopeContactableTime(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim sStaffCode
	Dim sContactableTime	: sContactableTime = ""

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	sSQL = "sp_GetDataContactableTime '" & sStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		sContactableTime = sContactableTime & oRS.Collect("ContactableTimeName")
		oRS.MoveNext
		If GetRSState(oRS) = True Then sContactableTime = sContactableTime & "�@"
	Loop
	Call RSClose(oRS)

	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">��]�A�����ԑ�</th>"
	Response.Write "<td>"
	Response.Write sContactableTime
	Response.Write "</td>"
	Response.Write "</tr>"

End Function

'******************************************************************************
'�T�@�v�F�v���t�B�[���y�[�W�̊�]�����Ǝ핔�����o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailStaff �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���O�C�������[�U���
'�@�@�@�FvUserID		�F���O�C�������[�U�h�c
'�g�p���F
'���@�l�F
'�X�@�V�F2018/07/10 LIS takaya �쐬
'******************************************************************************
Function DspProfileSideJob(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sStaffCode
    Dim sMsg

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">����</th>"
	Response.Write "<td>"

	sSQL = "select doit,AtHome,Attendance from p_SideJob where staffcode = '" & sStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	if GetRSState(oRS) = True then
        if oRS.Collect("Doit") = "2" then
            sMsg = "��]����"

            if oRS.Collect("AtHome") = "1" or oRS.Collect("Attendance") = "1"  then
                sMsg = sMsg + "�@( " 

                if oRS.Collect("AtHome") = "1" then
                    sMsg = sMsg + " �ݑ� "
                end if

                if oRS.Collect("Attendance") = "1" then
                    sMsg = sMsg + " �o�� "
                end if

                sMsg = sMsg + " )" 
            end if

            Response.Write sMsg 
        else
            Response.Write "��]���Ȃ�"
        end if
    else
        Response.Write "��]���Ȃ�"
	end if
	Call RSClose(oRS)

	Response.Write "</td>"
	Response.Write "</tr>"
End Function

%>
