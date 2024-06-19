<%
'******************************************************************************
'名　称：clsN_Image
'概　要：formで飛んできたN_Imageテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/10/06
'更　新：
'******************************************************************************
Class clsN_Image
	Public StaffCode
	Public CategoryCode
	Public Code
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsN_Image クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/10/06
	'更　新：
	'******************************************************************************
	Public Sub Initialize(ByVal vCategoryCode)
		Dim sTemp

		CategoryCode = vCategoryCode
		If Request.Form("conf_" & CategoryCode) <> "" Then sTemp = Replace(Request.Form("conf_" & CategoryCode), " ", "")

		IsData = False
		If sTemp <> "" Then
			Code = Split(sTemp, ",")
			MaxIndex = UBound(Code)
			IsData = True
		Else
			MaxIndex = -1
		End If

		Err = ""
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：up_Reg_N_Image 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC up_Del_N_Image '" & ChkSQLStr(vStaffCode) & "', '" & ChkSQLStr(CategoryCode) & "'" & vbCrLf
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC up_Reg_N_Image" & _
				" '" & ChkSQLStr(vStaffCode) & "'" & _
				",'" & ChkSQLStr(CategoryCode) & "'" & _
				",'" & ChkSQLStr(Code(idx)) & "'" & vbCrLf
		Next
	End Function
End Class
%>
