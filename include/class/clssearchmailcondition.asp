<%
'******************************************************************************
'�T�@�v�F���[���ꗗ�̌���������ێ�����N���X
'�ց@���F��Private
'�@�@�@�F
'�@�@�@�F��Public
'�@�@�@�FClass_Initialize	�F�R���X�g���N�^
'�@�@�@�FGetSearchParam		�F���[���ꗗ�̂f�d�s�p�����[�^�擾
'�@�@�@�FGetSQLSearchMail	�F���[���ꗗ�r�p�k�擾
'�@�@�@�F
'���@�l�F������ �����p�p�����[�^ (�A�h�z�b�N�Ȃr�p�k����)
'�@�@�@�FDayFrom				�F���t����[YYYYMMDD]
'�@�@�@�FDayTo					�F���t���[YYYYMMDD]
'�@�@�@�FOrderCode				�F���R�[�h
'�@�@�@�FSearchCode				�F����R�[�h
'�@�@�@�FEvaluation				�F�]��
'�@�@�@�FKeyword				�F�L�[���[�h
'�@�@�@�FMailContactPersonName	�F���l�S����
'�@�@�@�F
'�X�@�V�F2007/12/25 LIS K.Kokubo �쐬
'�@�@�@�F2009/03/27 LIS K.Kokubo ���C MailHistory.RegistDay�폜��SendDay�œ���
'�@�@�@�F2009/07/30 LIS K.Kokubo �X�J�E�g�A�v���[�`�t���O�����ǉ�
'�@�@�@�F2010/05/13 LIS K.Kokubo ���ǃ��[�������ǉ�(��M���̂�)
'******************************************************************************
Class clsSearchMailCondition
	Public UserCode

	'�������������o�ϐ�
	Public DayFrom
	Public DayTo
	Public OrderCode
	Public SearchCode
	Public Evaluation
	Public Keyword
	Public MailContactPersonName
	Public ScoutApproachFlag
	Public NotOpenFlag

	'���̑������o�ϐ�
	Public HtmlStaffSearch	'���������o�͂g�s�l�k��
	Public SQLStaffSearch	'�����r�p�k
	Public SQLWriteLog		'���O�������݂r�p�k

	'******************************************************************************
	'�T�@�v�F�R���X�g���N�^
	'���@���F
	'���@�l�F
	'�X�@�V�F2007/12/26 LIS K.Kokubo �쐬
	'******************************************************************************
	Private Sub Class_Initialize()
		UserCode = Session("userid")
		'�p�����[�^���猟���������擾
		If GetForm("sdf", 2) <> "" Then DayFrom = GetForm("sdf", 2)
		If GetForm("sdt", 2) <> "" Then DayTo = GetForm("sdt", 2)
		If GetForm("soc", 2) <> "" Then OrderCode = GetForm("soc", 2)
		If GetForm("sc", 2) <> "" Then SearchCode = GetForm("sc", 2)
		If GetForm("se", 2) <> "" Then Evaluation = GetForm("se", 2)
		If GetForm("skwd", 2) <> "" Then Keyword = GetForm("skwd", 2)
		If GetForm("mcpn", 2) <> "" Then MailContactPersonName = GetForm("mcpn", 2)
		If GetForm("ssaf", 2) <> "" Then ScoutApproachFlag = GetForm("ssaf", 2)
		If GetForm("snof", 2) <> "" Then NotOpenFlag = GetForm("snof", 2)
	End Sub

	'******************************************************************************
	'�T�@�v�F���[���ꗗ��GET�p�����[�^�𐶐����Ď擾�B
	'���@�l�F������
	'�@�@�@�F�p�����[�^���܂�URL�́AIE�̐�����2048�����܂łł���̂ŁA����ɍ��킹��B
	'�X�@�V�F2007/12/26 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Function GetSearchParam()
		GetSearchParam = ""
		If DayFrom <> "" Then GetSearchParam = GetSearchParam & "&amp;sdf=" & DayFrom
		If DayTo <> "" Then GetSearchParam = GetSearchParam & "&amp;sdt=" & DayTo
		If OrderCode <> "" Then GetSearchParam = GetSearchParam & "&amp;soc=" & OrderCode
		If SearchCode <> "" Then GetSearchParam = GetSearchParam & "&amp;sc=" & SearchCode
		If Evaluation <> "" Then GetSearchParam = GetSearchParam & "&amp;se=" & Evaluation
		If Keyword <> "" Then GetSearchParam = GetSearchParam & "&amp;skwd=" & Server.URLEncode(Keyword)
		If MailContactPersonName <> "" Then GetSearchParam = GetSearchParam & "&amp;mcpn=" & Server.URLEncode(MailContactPersonName)
		If ScoutApproachFlag <> "" Then GetSearchParam = GetSearchParam & "&amp;ssaf=" & ScoutApproachFlag
		If NotOpenFlag <> "" Then GetSearchParam = GetSearchParam & "&amp;snof=" & NotOpenFlag

		If GetSearchParam <> "" Then
			'����&amp;���폜
			GetSearchParam = Mid(GetSearchParam, 6)

			'�h�d�̎d�l�̓p�����[�^�̏�����Q�O�S�W�o�C�g
			GetSearchParam = Left(GetSearchParam, 2048)
		End If
	End Function

	'******************************************************************************
	'�T�@�v�F���[���ꗗ�r�p�k�擾
	'���@���FvMode	�F����M�t���O	["1"]���M���[�h [<>"1"]��M���[�h
	'���@�l�F
	'�X�@�V�F2007/12/25 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Function GetSQLSearchMail(ByVal vMode)
		Dim sSQL
		Dim tmpSQL1
		Dim tmpSQL2
		Dim sDeclare
		Dim sParams
		Dim sWhere
		Dim sJoin
		Dim idx

		sDeclare = ""
		sParams = ""
		sWhere = ""
		sJoin = ""

		tmpSQL1 = ""
		tmpSQL2 = ""

		'���O�C�������[�U
		sDeclare = sDeclare & "@vUserCode VARCHAR(8)"
		sParams = sParams & ",@vUserCode = N'" & UserCode & "'"

		If vMode = "1" Then
			'���M���[���ꗗ

			'���Ԏw��
			tmpSQL1 = ""
			If DayFrom & DayTo <> "" Then
				If DayFrom <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vDayFrom VARCHAR(8)"
					sParams = sParams & ",@vDayFrom = N'" & DayFrom & "'"

					If tmpSQL1 <> "" Then tmpSQL1 = tmpSQL1 & "AND "
					tmpSQL1 = tmpSQL1 & "A.SendDay >= CONVERT(DATETIME, @vDayFrom) "
				End If

				If DayTo <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vDayTo VARCHAR(8)"
					sParams = sParams & ",@vDayTo = N'" & DayTo & "'"

					If tmpSQL1 <> "" Then tmpSQL1 = tmpSQL1 & "AND "
					tmpSQL1 = tmpSQL1 & "A.SendDay < DATEADD(DAY, 1, CONVERT(DATETIME, @vDayTo)) "
				End If

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.SenderCode = @vUserCode AND " & tmpSQL1 & ") AS MDAY ON MH.ID = MDAY.ID "
			End If

			'���R�[�h
			If OrderCode <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vOrderCode VARCHAR(8)"
				sParams = sParams & ",@vOrderCode = N'" & OrderCode & "'"

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.SenderCode = @vUserCode AND A.OrderCode = @vOrderCode) AS MORD ON MH.ID = MORD.ID "
			End If

			'����R�[�h
			If SearchCode <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vSearchCode VARCHAR(8)"
				sParams = sParams & ",@vSearchCode = N'" & SearchCode & "'"

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.SenderCode = @vUserCode AND A.ReceiverCode = @vSearchCode) AS MSCD ON MH.ID = MSCD.ID "
			End If

			'�]��
			If Evaluation <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vEvaluation VARCHAR(8)"
				sParams = sParams & ",@vEvaluation = N'" & Evaluation & "'"

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.SenderCode = @vUserCode AND A.SenderEvaluation = @vEvaluation) AS MEVL ON MH.ID = MEVL.ID "
			End If

			'�L�[���[�h
			If Keyword <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vKeyword VARCHAR(100)"
				sParams = sParams & ",@vKeyword = N'" & Keyword & "'"

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.SenderCode = @vUserCode AND A.Subject + ':' + ISNULL(A.SenderRemark, '') + ':' + A.Body LIKE '%' + @vKeyword + '%') AS MWRD ON MH.ID = MWRD.ID "
			End If

			'���l�S����
			If MailContactPersonName <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vMailContactPersonName VARCHAR(100)"
				sParams = sParams & ",@vMailContactPersonName = N'" & MailContactPersonName & "'"

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A INNER JOIN C_Contact AS B ON A.OrderCode = B.OrderCode AND B.PersonName = @vMailContactPersonName WHERE A.SenderCode = @vUserCode) AS MCPN ON MH.ID = MCPN.ID "
			End If

			'�X�J�E�g�A�v���[�`�t���O
			If ScoutApproachFlag <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vScoutApproachFlag VARCHAR(1)"
				sParams = sParams & ",@vScoutApproachFlag = N'" & ScoutApproachFlag & "'"

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.SenderCode = @vUserCode AND A.ScoutApproachFlag = @vScoutApproachFlag) AS SAF ON MH.ID = SAF.ID "
			End If

			sSQL = ""
			sSQL = sSQL & "SELECT MH.ID, MH.SendDay "
			sSQL = sSQL & "FROM MailHistory AS MH " & sJoin
			sSQL = sSQL & "WHERE MH.SenderCode = @vUserCode "
			sSQL = sSQL & "AND MH.SenderDelFlag = '0' "
			sSQL = sSQL & "OPTION(MAXDOP 1) "

			'�p�����[�^�N�G����
			sSQL = "" & _
			"/*�i�r�E���M���[���ꗗ*/ " & vbCrLf & _
			"SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED " & vbCrLf & _
			"EXEC sp_executesql N'" & Replace(sSQL, "'", "''") & "'"
			If sDeclare <> "" Then sSQL = sSQL & ",N'" & sDeclare & "'" & sParams
		Else
			'��M���[���ꗗ

			'���Ԏw��
			tmpSQL1 = ""
			If DayFrom & DayTo <> "" Then
				If DayFrom <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vDayFrom VARCHAR(8)"
					sParams = sParams & ",@vDayFrom = N'" & DayFrom & "'"

					If tmpSQL1 <> "" Then tmpSQL1 = tmpSQL1 & "AND "
					tmpSQL1 = tmpSQL1 & "A.SendDay >= CONVERT(DATETIME, @vDayFrom) "
				End If

				If DayTo <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vDayTo VARCHAR(8)"
					sParams = sParams & ",@vDayTo = N'" & DayTo & "'"

					If tmpSQL1 <> "" Then tmpSQL1 = tmpSQL1 & "AND "
					tmpSQL1 = tmpSQL1 & "A.SendDay < DATEADD(DAY, 1, CONVERT(DATETIME, @vDayTo)) "
				End If

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.ReceiverCode = @vUserCode AND " & tmpSQL1 & ") AS MDAY ON MH.ID = MDAY.ID "
			End If

			'���R�[�h
			If OrderCode <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vOrderCode VARCHAR(8)"
				sParams = sParams & ",@vOrderCode = N'" & OrderCode & "'"

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.ReceiverCode = @vUserCode AND A.OrderCode = @vOrderCode) AS MORD ON MH.ID = MORD.ID "
			End If

			'����R�[�h
			If SearchCode <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vSearchCode VARCHAR(8)"
				sParams = sParams & ",@vSearchCode = N'" & SearchCode & "'"

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.ReceiverCode = @vUserCode AND A.SenderCode = @vSearchCode) AS MSCD ON MH.ID = MSCD.ID "
			End If

			'�]��
			If Evaluation <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vEvaluation VARCHAR(8)"
				sParams = sParams & ",@vEvaluation = N'" & Evaluation & "'"

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.ReceiverCode = @vUserCode AND A.ReceiverEvaluation = @vEvaluation) AS MEVL ON MH.ID = MEVL.ID "
			End If

			'�L�[���[�h
			If Keyword <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vKeyword VARCHAR(100)"
				sParams = sParams & ",@vKeyword = N'" & Keyword & "'"

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.ReceiverCode = @vUserCode AND A.Subject + ':' + ISNULL(A.ReceiverRemark, '') + ':' + A.Body LIKE '%' + @vKeyword + '%') AS MWRD ON MH.ID = MWRD.ID "
			End If

			'���l�S����
			If MailContactPersonName <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vMailContactPersonName VARCHAR(100)"
				sParams = sParams & ",@vMailContactPersonName = N'" & MailContactPersonName & "'"

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A INNER JOIN C_Contact AS B ON A.OrderCode = B.OrderCode AND B.PersonName = @vMailContactPersonName WHERE A.ReceiverCode = @vUserCode) AS MCPN ON MH.ID = MCPN.ID "
			End If

			'�X�J�E�g�A�v���[�`�t���O
			If ScoutApproachFlag <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vScoutApproachFlag VARCHAR(1)"
				sParams = sParams & ",@vScoutApproachFlag = N'" & ScoutApproachFlag & "'"

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.ReceiverCode = @vUserCode AND A.ScoutApproachFlag = @vScoutApproachFlag) AS SAF ON MH.ID = SAF.ID "
			End If

			'���ǃ��[��
			If NotOpenFlag <> "" Then
				If NotOpenFlag = "1" Then
					'����
					sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.ReceiverCode = @vUserCode AND A.OpenDay IS NULL) AS NRF ON MH.ID = NRF.ID "
				ElseIf NotOpenFlag = "0" Then
					'����
					sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.ReceiverCode = @vUserCode AND A.OpenDay > 0) AS NRF ON MH.ID = NRF.ID "
				End If
			End If

			sSQL = ""
			sSQL = sSQL & "SELECT MH.ID, MH.SendDay "
			sSQL = sSQL & "FROM MailHistory AS MH " & sJoin
			sSQL = sSQL & "WHERE MH.ReceiverCode = @vUserCode "
			sSQL = sSQL & "AND MH.ReceiverDelFlag = '0' "
			sSQL = sSQL & "OPTION(MAXDOP 1) "

			'�p�����[�^�N�G����
			sSQL = "" & _
			"/*�i�r�E��M���[���ꗗ*/ " & vbCrLf & _
			"SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED " & vbCrLf & _
			"EXEC sp_executesql N'" & Replace(sSQL, "'", "''") & "'"
			If sDeclare <> "" Then sSQL = sSQL & ",N'" & sDeclare & "'" & sParams
		End If

		GetSQLSearchMail = sSQL

'Response.Write GetSQLSearchMail
	End Function
End Class
%>
