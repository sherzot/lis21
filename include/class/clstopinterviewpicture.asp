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
'���@�́FclsPicture
'�T�@�v�F�o�C�i���`����form�Ŕ��ł���CompanyInfo�e�[�u���p�̃f�[�^�������߂̃N���X
'���@�l�F���O��dbconn���J���Ă���
'�쐬�ҁFLis Kokubo
'�쐬���F2006/05/19
'�X�@�V�F
'******************************************************************************
Class clsTopInterviewPicture
	Public OptionNo
	Public Picture
	Public Caption
	Public Size
	Public Mode
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'���@�́FInitialize
	'�T�@�v�FclsTopInterviewPicture�N���X�̏������֐�
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/05/19
	'�X�@�V�F
	'******************************************************************************
	Public Sub Initialize(ByVal vBinary)
		Dim oBasp	: Set oBasp = Server.CreateObject("basp21")

		Picture = oBasp.FormBinary(vBinary, "frmpicturepath")
		Mode = oBasp.Form(vBinary, "frmregflag")
		Size = oBasp.FormFileSize(vBinary, "frmpicturepath")

		IsData = False
		If (Mode = "1" And Size > 0) Or Mode = "0" Then IsData = True

		'�l�`�F�b�N
		Err = ""

		Set oBasp = Nothing
	End Sub

	'******************************************************************************
	'���@�́FReg
	'�T�@�v�F�摜��DB�ɓo�^
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/05/19
	'�X�@�V�F
	'******************************************************************************
	Public Function Reg()
		Dim sSQL
		Dim oRS
		Dim flgQE
		Dim sError
		Dim oBasp

		Set oBasp = Server.CreateObject("basp21")

		Reg = False
		If IsData = True And G_USERID <> "" And Size <= 51200 Then
			Set oRS = Server.CreateObject("ADODB.Recordset")
			oRS.CursorType = 2
			oRS.LockType = 3

			sSQL = ""
			sSQL = sSQL & "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
			sSQL = sSQL & "SELECT * FROM CMPTopInterview_Picture WHERE CompanyCode = '" & G_USERID & "'"
			oRS.Open sSQL, dbconn

			If oRS.EOF Then
				oRS.AddNew
				oRS.Fields("CompanyCode") = G_USERID
				oRS.Fields("RegistDay") = Now
			End If
			oRS.Fields("Picture").AppendChunk Picture
			oRS.Fields("UpdateDay") = Now

			oRS.Update
			oRS.Close

			Set oRS = Nothing
			Reg = True
		End If
		Set oBasp = Nothing
	End Function

	'******************************************************************************
	'���@�́FDel
	'�T�@�v�F�摜��DB�ɓo�^
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/05/19
	'�X�@�V�F
	'******************************************************************************
	Public Function Del()
		Dim sSQL
		Dim oRS
		Dim flgQE
		Dim sError

		sSQL = "EXEC up_DelCMPTopInterview_Picture '" & G_USERID & "'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

		Del = flgQE
	End Function
End Class
%>
