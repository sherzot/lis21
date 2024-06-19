<%
'******************************************************************************
'概　要：formで飛んできた画像ファイルをデータベースに登録するためのクラス
'備　考：事前に connect.asp をインクルードしておくこと！
'作成者：Lis Kokubo
'作成日：2006/05/19
'更　新：
'******************************************************************************
%>
<%
'******************************************************************************
'名　称：clsPicture
'概　要：バイナリ形式のformで飛んできたCompanyInfoテーブル用のデータを持つためのクラス
'備　考：事前にdbconnを開いておく
'作成者：Lis Kokubo
'作成日：2006/05/19
'更　新：
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
	'名　称：Initialize
	'概　要：clsTopInterviewPictureクラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/05/19
	'更　新：
	'******************************************************************************
	Public Sub Initialize(ByVal vBinary)
		Dim oBasp	: Set oBasp = Server.CreateObject("basp21")

		Picture = oBasp.FormBinary(vBinary, "frmpicturepath")
		Mode = oBasp.Form(vBinary, "frmregflag")
		Size = oBasp.FormFileSize(vBinary, "frmpicturepath")

		IsData = False
		If (Mode = "1" And Size > 0) Or Mode = "0" Then IsData = True

		'値チェック
		Err = ""

		Set oBasp = Nothing
	End Sub

	'******************************************************************************
	'名　称：Reg
	'概　要：画像をDBに登録
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/05/19
	'更　新：
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
	'名　称：Del
	'概　要：画像をDBに登録
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/05/19
	'更　新：
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
