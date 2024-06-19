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
'名　称：clsImage
'概　要：バイナリ形式のformで飛んできたCompanyInfoテーブル用のデータを持つためのクラス
'備　考：事前にdbconnを開いておく
'作成者：Lis Kokubo
'作成日：2006/05/19
'更　新：
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
	'概　要：コンストラクタ
	'備　考：
	'履　歴：2010/03/30 LIS K.Kokubo
	'******************************************************************************
	Private Sub Class_Initialize()
		UserID = Session("userid")
	End Sub

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsImageクラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/05/19
	'更　新：
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

		'値チェック
		Err = ""

		Set oBasp = Nothing
	End Sub

	'******************************************************************************
	'名　称：RegImage
	'概　要：画像をDBに登録
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/05/19
	'更　新：
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
	'名　称：RegOrderImage
	'概　要：求人票画像をDBに登録
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/05/19
	'更　新：
	'******************************************************************************
	Public Function RegOrderImage()
		Dim sSQL
		Dim oRS
		Dim oBasp	: Set oBasp = Server.CreateObject("basp21")

		Dim iNewOptionNo

		RegOrderImage = False

		'エラーチェック
		If Size > 204800 Then Err = Err & "写真のファイルサイズが大きすぎます。２００ＫＢ以下にして下さい。<br>"

		If IsData = True And UserID <> "" And Err = "" Then
			Set oRS = Server.CreateObject("ADODB.Recordset")
			oRS.CursorType = 2
			oRS.LockType = 3

			'******************************************************************************
			'OptionNo取得 start
			'******************************************************************************
			sSQL = "up_GetNewOptionNo '" & UserID & "'"
			Call oRS.Open(sSQL, dbconn)
			If GetRSState(oRS) = True Then
				iNewOptionNo = oRS.Collect("OptionNo")
			End If
			oRS.Close

			If Len(OptionNo) = 0 Then OptionNo = iNewOptionNo
			'******************************************************************************
			'OptionNo取得 end
			'******************************************************************************

			sSQL = "SELECT * FROM OptionPicture WHERE CompanyCode = '" & UserID & "' AND OptionNo = " & OptionNo
			Call oRS.Open(sSQL, dbconn)

			If GetRSState(oRS) = False Then
				'新規作成
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
				'更新
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
	'名　称：RegAdvertImage
	'概　要：画像をDBに登録
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/05/19
	'更　新：
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
	'名　称：RegResumeImage
	'概　要：画像をDBに登録
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/05/19
	'更　新：
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
	'名　称：DelImage
	'概　要：画像をDBに登録
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/05/19
	'更　新：
	'******************************************************************************
	Public Function DelImage()
		Dim sSQL
		Dim oRS

		sSQL = "sp_Del_Picture '" & UserID & "', '" & OptionNo & "'"
		dbconn.Execute(sSQL)
	End Function

	'******************************************************************************
	'名　称：DelAdvertImage
	'概　要：画像をDBに登録
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/05/19
	'更　新：
	'******************************************************************************
	Public Function DelAdvertImage(vAdvertCode)
		Dim sSQL
		Dim oRS

		sSQL = "sp_Del_Picture '" & vAdvertCode & "', '1'"
		dbconn.Execute(sSQL)
	End Function

	'******************************************************************************
	'名　称：DelResumeImage
	'概　要：画像をDBに登録
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/05/19
	'更　新：
	'******************************************************************************
	Public Function DelResumeImage(vStaffCode)
		Dim sSQL
		Dim oRS

		sSQL = "sp_Del_Picture '" & vStaffCode & "', '1'"
		dbconn.Execute(sSQL)
	End Function
End Class
%>
