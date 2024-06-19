<%
'******************************************************************************
'�T�@�v�Fform�Ŕ��ł����摜�t�@�C�����f�[�^�x�[�X�ɓo�^���邽�߂̃N���X
'���@�l�F���O�� connect.asp ���C���N���[�h���Ă������ƁI
'�쐬�ҁFLis Kokubo
'�쐬���F2006/05/19
'�X�@�V�F
'******************************************************************************
%>
<%
'******************************************************************************
'���@�́FclsImage
'�T�@�v�F�o�C�i���`����form�Ŕ��ł���CompanyInfo�e�[�u���p�̃f�[�^�������߂̃N���X
'���@�l�F���O��dbconn���J���Ă���
'�쐬�ҁFLis Kokubo
'�쐬���F2006/05/19
'�X�@�V�F
'******************************************************************************
Class clsImage
	Public UserID
	Public OptionNo
	Public Image
	Public Caption
	Public Size
	Public Mode
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'�T�@�v�F�R���X�g���N�^
	'���@�l�F
	'���@���F2010/03/30 LIS K.Kokubo
	'******************************************************************************
	Private Sub Class_Initialize()
		UserID = Session("userid")
	End Sub

	'******************************************************************************
	'���@�́FInitialize
	'�T�@�v�FclsImage�N���X�̏������֐�
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/05/19
	'�X�@�V�F
	'******************************************************************************
	Public Sub Initialize(ByVal vBinary, ByVal vImageInputName)
		Dim oBasp	: Set oBasp = Server.CreateObject("basp21")
		MaxIndex = -1

		OptionNo = oBasp.Form(vBinary, "CONF_OptionNo")
		Image = oBasp.FormBinary(vBinary, vImageInputName)
		Caption = oBasp.Form(vBinary, "CONF_Caption")
		Mode = oBasp.Form(vBinary, "CONF_ImageFlag")
		Size = oBasp.FormFileSize(vBinary, vImageInputName)

		IsData = False
		If (Mode = "1" And (Size > 0 Or Caption <> "")) Or Mode = "0" Then IsData = True

		'�l�`�F�b�N
		Err = ""

		Set oBasp = Nothing
	End Sub

	'******************************************************************************
	'���@�́FRegImage
	'�T�@�v�F�摜��DB�ɓo�^
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/05/19
	'�X�@�V�F
	'******************************************************************************
	Public Function RegImage()
		Dim sSQL
		Dim oRS
		Dim oBasp	: Set oBasp = Server.CreateObject("basp21")

		RegImage = False
		If IsData = True And UserID <> "" And Size <= 204800 Then
			Set oRS = Server.CreateObject("ADODB.Recordset")
			oRS.CursorType = 2
			oRS.LockType = 3

			sSQL = "SELECT * FROM OptionPicture WHERE CompanyCode = '" & UserID & "' AND OptionNo = 1"
			oRS.Open sSQL, dbconn

			If oRS.EOF Then
				oRS.AddNew
				oRS.Fields("CompanyCode") = UserID
				oRS.Fields("OptionNo") = 1
				oRS.Fields("RegistDate") = Now
			End If

			oRS.Fields("Picture").AppendChunk Image
			oRS.Fields("UpdateDate") = Now

			oRS.Update
			oRS.Close

			Set oRS = Nothing
			RegImage = True
		End If
		Set oBasp = Nothing
	End Function

	'******************************************************************************
	'���@�́FRegOrderImage
	'�T�@�v�F���l�[�摜��DB�ɓo�^
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/05/19
	'�X�@�V�F
	'******************************************************************************
	Public Function RegOrderImage()
		Dim sSQL
		Dim oRS
		Dim oBasp	: Set oBasp = Server.CreateObject("basp21")

		Dim iNewOptionNo

		RegOrderImage = False

		'�G���[�`�F�b�N
		If Size > 204800 Then Err = Err & "�ʐ^�̃t�@�C���T�C�Y���傫�����܂��B�Q�O�O�j�a�ȉ��ɂ��ĉ������B<br>"

		If IsData = True And UserID <> "" And Err = "" Then
			Set oRS = Server.CreateObject("ADODB.Recordset")
			oRS.CursorType = 2
			oRS.LockType = 3

			'******************************************************************************
			'OptionNo�擾 start
			'******************************************************************************
			sSQL = "up_GetNewOptionNo '" & UserID & "'"
			Call oRS.Open(sSQL, dbconn)
			If GetRSState(oRS) = True Then
				iNewOptionNo = oRS.Collect("OptionNo")
			End If
			oRS.Close

			If Len(OptionNo) = 0 Then OptionNo = iNewOptionNo
			'******************************************************************************
			'OptionNo�擾 end
			'******************************************************************************

			sSQL = "SELECT * FROM OptionPicture WHERE CompanyCode = '" & UserID & "' AND OptionNo = " & OptionNo
			Call oRS.Open(sSQL, dbconn)

			If GetRSState(oRS) = False Then
				'�V�K�쐬
				If Size > 0 Then
					oRS.AddNew
					oRS.Fields("CompanyCode") = UserID
					oRS.Fields("OptionNo") = OptionNo
					oRS.Fields("RegistDate") = Now
					oRS.Fields("Picture").AppendChunk Image
					If Caption <> "" Then oRS.Fields("Caption") = Caption
					oRS.Fields("UpdateDate") = Now
					oRS.Update
				End If
			Else
				'�X�V
				oRS.Fields("OptionNo") = OptionNo
				If Size > 0 Then oRS.Fields("Picture").AppendChunk Image
				If Caption <> "" Then oRS.Fields("Caption") = Caption
				oRS.Fields("UpdateDate") = Now
				oRS.Update
			End If

			Call RSClose(oRS)

			RegOrderImage = True
		End If
		Set oBasp = Nothing
	End Function

	'******************************************************************************
	'���@�́FRegAdvertImage
	'�T�@�v�F�摜��DB�ɓo�^
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/05/19
	'�X�@�V�F
	'******************************************************************************
	Public Function RegAdvertImage(vAdvertCode)
		Dim sSQL
		Dim oRS
		Dim oBasp	: Set oBasp = Server.CreateObject("basp21")

		RegAdvertImage = False
		If IsData = True And UserID <> "" And Size <= 102400 Then
			Set oRS = Server.CreateObject("ADODB.Recordset")
			oRS.CursorType = 2
			oRS.LockType = 3

			sSQL = "SELECT * FROM C_AdvertInfo WHERE AdvertCode = '" & vAdvertCode & "'"
			oRS.Open sSQL, dbconn

			oRS.Fields("Picture").AppendChunk Image
			oRS.Fields("UpdateDay") = Now

			oRS.Update
			oRS.Close

			Set oRS = Nothing
			RegAdvertImage = True
		End If
		Set oBasp = Nothing
	End Function

	'******************************************************************************
	'���@�́FRegResumeImage
	'�T�@�v�F�摜��DB�ɓo�^
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/05/19
	'�X�@�V�F
	'******************************************************************************
	Public Function RegResumeImage(vStaffCode)
		Dim sSQL
		Dim oRS
		Dim oBasp: Set oBasp = Server.CreateObject("basp21")

		RegResumeImage = False
		If IsData = True And UserID <> "" And Size <= 512000 Then
			Set oRS = Server.CreateObject("ADODB.Recordset")
			oRS.CursorType = 2
			oRS.LockType = 3

			sSQL = "SELECT * FROM P_Picture WHERE StaffCode = '" & vStaffCode & "'"
			oRS.Open sSQL, dbconn

			If oRS.EOF Then
				oRS.AddNew
				oRS.Fields("StaffCode") = UserID
				oRS.Fields("RegistDay") = Now
			End If

			oRS.Fields("Picture").AppendChunk Image
			oRS.Fields("UpdateDay") = Now

			oRS.Update
			oRS.Close

			Set oRS = Nothing
			RegResumeImage = True
		End If
		Set oBasp = Nothing
	End Function

	'******************************************************************************
	'���@�́FDelImage
	'�T�@�v�F�摜��DB�ɓo�^
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/05/19
	'�X�@�V�F
	'******************************************************************************
	Public Function DelImage()
		Dim sSQL
		Dim oRS

		sSQL = "sp_Del_Picture '" & UserID & "', '" & OptionNo & "'"
		dbconn.Execute(sSQL)
	End Function

	'******************************************************************************
	'���@�́FDelAdvertImage
	'�T�@�v�F�摜��DB�ɓo�^
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/05/19
	'�X�@�V�F
	'******************************************************************************
	Public Function DelAdvertImage(vAdvertCode)
		Dim sSQL
		Dim oRS

		sSQL = "sp_Del_Picture '" & vAdvertCode & "', '1'"
		dbconn.Execute(sSQL)
	End Function

	'******************************************************************************
	'���@�́FDelResumeImage
	'�T�@�v�F�摜��DB�ɓo�^
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/05/19
	'�X�@�V�F
	'******************************************************************************
	Public Function DelResumeImage(vStaffCode)
		Dim sSQL
		Dim oRS

		sSQL = "sp_Del_Picture '" & vStaffCode & "', '1'"
		dbconn.Execute(sSQL)
	End Function
End Class
%>
