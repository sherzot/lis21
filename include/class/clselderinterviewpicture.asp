<%
'******************************************************************************
'概　要：formで飛んできた画像ファイルをデータベースに登録するためのクラス
'備　考：事前に connect.asp をインクルードしておくこと！
'更　新：2008/01/31 LIS K.Kokubo 作成
'******************************************************************************
%>
<%
'******************************************************************************
'名　称：clsElderInterviewPicture
'概　要：バイナリ形式のformで飛んできたCompanyInfoテーブル用のデータを持つためのクラス
'備　考：事前にdbconnを開いておく
'更　新：2008/01/31 LIS K.Kokubo 作成
'******************************************************************************
Class clsElderInterviewPicture
	Public OrderCode
	Public Seq
	Public Picture
	Public Size
	Public Mode
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsElderInterviewPictureクラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/05/19
	'更　新：
	'******************************************************************************
	Public Sub Initialize(ByVal vBinary)
		Dim oBasp	: Set oBasp = Server.CreateObject("basp21")

		OrderCode = oBasp.Form(vBinary, "frmordercode")
		Seq = oBasp.Form(vBinary, "frmseq")
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
		Dim oRS2
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
			sSQL = sSQL & "SELECT * FROM C_ElderInterview_Picture WHERE OrderCode = '" & OrderCode & "' AND Seq = '" & Seq & "'"
			oRS.Open sSQL, dbconn

			If oRS.EOF Then
				oRS.AddNew
				oRS.Fields("OrderCode") = OrderCode
				oRS.Fields("Seq") = Seq
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

		sSQL = "EXEC up_DelC_ElderInterview_Picture '" & OrderCode & "', '" & Seq & "'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

		Del = flgQE
	End Function
End Class
%>
