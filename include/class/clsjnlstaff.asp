<%
'******************************************************************************
'�T�@�v�F�k�h�r�W���[�i���̋��E�Ҕz�M�\����e�N���X(�\����)
'�ց@���F��Private
'�@�@�@�F��Public
'���@�l�F
'�X�@�V�F2009/01/14 LIS K.Kokubo �쐬
'�@�@�@�F2009/02/17 LIS K.Kokubo �ǉ� [HopeJobType][HopeYearlyIncome][EducateHistory]
'******************************************************************************
Class clsJNLStaff
	Public BranchCode
	Public ContentsType
	Public vol
	Public StaffCode
	Public Age
	Public Sex
	Public Address
	Public WorkStartDay
	Public CareerHistory
	Public CounselingView
	Public Skill
	Public RecentConditions
	Public Hope
	Public HopeJobType
	Public HopeYearlyIncome
	Public EducateHistory
	Public Certify

	'******************************************************************************
	'�T�@�v�F�����o�ϐ�������
	'���@���F
	'���@�l�F
	'�g�@�p�F�Г�/mailservice/lisjournal/reserve.asp
	'�X�@�V�F2009/01/14 LIS K.Kokubo �쐬
	'�@�@�@�F2009/02/17 LIS K.Kokubo �ǉ� [HopeJobType][HopeYearlyIncome][EducateHistory]
	'******************************************************************************
	Public Function Clear()
		BranchCode = Empty
		ContentsType = Empty
		vol = Empty
		StaffCode = Empty
		Age = Empty
		Sex = Empty
		Address = Empty
		WorkStartDay = Empty
		CareerHistory = Empty
		CounselingView = Empty
		Skill = Empty
		RecentConditions = Empty
		Hope = Empty
		HopeJobType = Empty
		HopeYearlyIncome = Empty
		EducateHistory = Empty
		Certify = Empty
	End Function

	'******************************************************************************
	'�T�@�v�F�k�h�r�W���[�i���̋Ζ��J�n�\����ϊ�
	'���@���F
	'���@�l�F
	'�g�@�p�F�Г�/mailservice/lisjournal/reserve.asp
	'�X�@�V�F2007/09/10 LIS K.Kokubo �쐬
	'�@�@�@�F2008/04/24 LIS K.Kokubo ���x�Ђ����u���Г������k�v����
	'�@�@�@�F2009/01/14 LIS K.Kokubo �N���X�p�ɕύX
	'******************************************************************************
	Public Function ChgWorkStartDay()
		On Error Resume Next
		Dim dWorkStartDay

		ChgWorkStartDay = ""

		If IsDate(WorkStartDay) = True Then
			dWorkStartDay = CDate(WorkStartDay)

			If DateDiff("d", dWorkStartDay, Date) < 0 Then
				ChgWorkStartDay = Year(dWorkStartDay) & "�N" & Month(dWorkStartDay) & "��" & Day(dWorkStartDay) & "��"
			Else
				ChgWorkStartDay = "�����\"
			End If
		ElseIf BranchCode = "OR" Then
			ChgWorkStartDay = "���Г������k"
		Else
			ChgWorkStartDay = "��������"
		End If
	End Function
End Class
%>
