<%
'******************************************************************************
'名　称：clsP_DevelopmentTool
'概　要：formで飛んできたP_DevelopmentToolテーブル用のデータを持つためのクラス
'　　　：CategoryCode毎にこのクラスを作成して使用する。
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_DevelopmentTool
	Public StaffCode
	Public CareerHistoryID()
	Public CategoryCode
	Public Code()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_DevelopmentTool クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize(vCategoryCode)
		Dim iCareerHistoryID
		Dim idx		: idx = 1
		Dim idx1
		Dim idx2
		Dim flag	: flag = False

		IsData = False
		MaxIndex = -1
		StaffCode = Request.Form("CONF_StaffCode")

		Err = ""

		'CONF_ の名前が Lang, App, DB と略されてしまっている事への処置
		Select Case vCategoryCode
			Case "Lang": CategoryCode = "DevelopmentLanguage"
			Case "App": CategoryCode = "Application"
			Case "DB": CategoryCode = "Database"
			Case Else: CategoryCode = vCategoryCode
		End Select

		iCareerHistoryID = 0
		Do While True
			If ExistsForm("CONF_DevelopmentTool_" & vCategoryCode & idx) = False Then Exit Do

			If Request.Form("CONF_DevelopmentDetail" & idx) <> "" Then
				iCareerHistoryID = iCareerHistoryID + 1
				flag = True
			End If

			If flag = True And Request.Form("CONF_DevelopmentTool_" & vCategoryCode & idx) <> "" Then
				MaxIndex = MaxIndex + 1
				ReDim Preserve CareerHistoryID(MaxIndex)
				ReDim Preserve Code(MaxIndex)
				If flag = True Then
					CareerHistoryID(MaxIndex) = iCareerHistoryID
					Code(MaxIndex) = Split(Request.Form("CONF_DevelopmentTool_" & vCategoryCode & idx), ",")
				End If
			End If

			idx = idx + 1
			flag = False
		Loop

		For idx1 = 0 To MaxIndex
			For idx2 = LBound(Code(idx1)) To UBound(Code(idx1))
				Code(idx1)(idx2) = Trim(Code(idx1)(idx2))
				If Code(idx1)(idx2) <> "" And IsNumber(Code(idx1)(idx2), 3, False) = True Then
					IsData = True
				Else
					Err = Err & "Code(" & idx1 & ")(" & idx2 & ")" & vbCrLf
				End If
			Next
		Next
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_DevelopmentTool 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx
		Dim idxCode

		GetRegSQL = "EXEC sp_Del_P_DevelopmentTool '" & ChkSQLStr(vStaffCode) & "', '" & ChkSQLStr(CategoryCode) & "'" & vbCrLf
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			For idxCode = LBound(Code(idx)) To UBound(Code(idx))
				GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_DevelopmentTool" & _
					" '" & ChkSQLStr(vStaffCode) & "'" & _
					",'" & ChkSQLStr(CareerHistoryID(idx)) & "'" & _
					",''" & _
					",'" & ChkSQLStr(CategoryCode) & "'" & _
					",'" & ChkSQLStr(Code(idx)(idxCode)) & "'" & vbCrLf
			Next
		Next
	End Function
End Class
%>
