<%
'**********************************************************************************************************************
'概　要：レポートで使用する関数群
'　　　：
'　　　：■■■　前提条件　■■■
'　　　：要事前インクルード
'　　　：/config/personel.asp
'　　　：/include/commonfunc.asp
'一　覧：■■■　メール一覧ページ出力用　■■■
'　　　：GetHtmlOrderList	：レポート（受注・登録者一覧）の受注部分のＨＴＭＬ取得
'　　　：GetHtmlStaffList	：レポート（受注・登録者一覧）の登録者部分のＨＴＭＬ取得
'**********************************************************************************************************************

'******************************************************************************
'概　要：レポート（受注・登録者一覧）の受注部分のＨＴＭＬ取得
'引　数：
'備　考：
'使　用：社内/report/order_staff_list.asp
'更　新：2007/09/20 LIS K.Kokubo 作成
'******************************************************************************
Function GetHtmlOrderList(ByRef rDB, ByVal vYM, ByVal vBranchCode)
	Dim sSQL
	Dim oRS
	Dim sError
	Dim flgQE

	Dim dbOrderCode
	Dim dbCompanyName
	Dim dbJobTypeName
	Dim dbJobTypeDetail
	Dim dbStaffName
	Dim dbSelectionTypeName

	Dim sHTML
	Dim sOldOrderCode
	Dim sOldCompany
	Dim sOldStaff
	Dim sStyle
	Dim iRow
	Dim iAbsPos
	Dim iFltRec
	Dim iFltRec2

	sOldOrderCode = ""
	sOldCompany = ""
	sOldStaff = ""

	sSQL = "up_LstRptOrder '" & vYM & "', '" & vBranchCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	sHTML = ""
	sHTML = sHTML & "<table class=""orderstafflist"" border=""0"" style=""width:380px;"">" & vbCrLf
	sHTML = sHTML & "<thead>" & vbCrLf
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<th colspan=""5"">受注</th>"
	sHTML = sHTML & "</tr>" & vbCrLf
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td>No.</td>"
	sHTML = sHTML & "<td>社名</td>"
	sHTML = sHTML & "<td>職種</td>"
	sHTML = sHTML & "<td>エントリー</td>"
	sHTML = sHTML & "<td>結果</td>"
	sHTML = sHTML & "</tr>" & vbCrLf
	sHTML = sHTML & "</thead>" & vbCrLf
	sHTML = sHTML & "<tbody>" & vbCrLf
	Do While GetRSState(oRS) = True
		sStyle = ""
		iFltRec = 1

		dbOrderCode = oRS.Collect("OrderCode")
		dbCompanyName = oRS.Collect("CompanyName")
		dbJobTypeName = Replace(oRS.Collect("JobTypeName"), vbCrLf, "<br>")
		dbJobTypeDetail = oRS.Collect("JobTypeDetail")
		dbStaffName = oRS.Collect("StaffName")
		dbSelectionTypeName = oRS.Collect("SelectionTypeName")

		'求職者毎の選考数を取得
		If sOldCompany <> dbCompanyName Then
			iAbsPos = oRS.AbsolutePosition
			oRS.Filter = "CompanyName = '" & dbCompanyName & "'"
			iFltRec = oRS.RecordCount
			oRS.Filter = 0
			oRS.AbsolutePosition = iAbsPos
		End If

		If sOldCompany <> dbCompanyName And sOldCompany <> "" Then
			sStyle = "border-top:1px dotted #666666;"
		End If

		If sOldCompany = dbCompanyName And sOldCompany <> "" Then
			dbCompanyName = ""
		End If

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td style=""" & sStyle & """>No.</td>"
		sHTML = sHTML & "<td style=""" & sStyle & """>" & dbCompanyName & "</td>"

		iFltRec2 = 0
		If sOldOrderCode <> dbOrderCode Then
			iAbsPos = oRS.AbsolutePosition
			oRS.Filter = "OrderCode = '" & dbOrderCode & "'"
			iFltRec2 = oRS.RecordCount
			oRS.Filter = 0
			oRS.AbsolutePosition = iAbsPos
		End If

		If iFltRec2 > 1 Then
			sHTML = sHTML & "<td rowspan=""" & iFltRec2 & """ style=""" & sStyle & " border-bottom:1px dotted #666666;"">" & dbJobTypeDetail & "</td>"
		ElseIf iFltRec2 = 1 Then
			sHTML = sHTML & "<td style=""" & sStyle & " border-bottom:1px dotted #666666;"">" & dbJobTypeDetail & "</td>"
		End If
		sHTML = sHTML & "<td style=""" & sStyle & " border-bottom:1px dotted #666666;"">" & dbStaffName & "</td>"
		sHTML = sHTML & "<td style=""" & sStyle & " border-bottom:1px dotted #666666;"">" & dbSelectionTypeName & "</td>"
		sHTML = sHTML & "</tr>" & vbCrLf

		sOldOrderCode = oRS.Collect("OrderCode")
		sOldCompany = oRS.Collect("CompanyName")
		oRS.MoveNext
	Loop
	sHTML = sHTML & "<tr class=""setwidth"">" & vbCrLf
	sHTML = sHTML & "<th style=""width:30px;""></th>"
	sHTML = sHTML & "<th style=""width:80px;""></th>"
	sHTML = sHTML & "<th style=""width:80px;""></th>"
	sHTML = sHTML & "<th style=""width:80px;""></th>"
	sHTML = sHTML & "<th style=""width:30px;""></th>"
	sHTML = sHTML & "</tr>" & vbCrLf
	sHTML = sHTML & "</tbody>" & vbCrLf
	sHTML = sHTML & "<tfoot>" & vbCrLf
	sHTML = sHTML & "<tr><th colspan=""5""></th></tr>" & vbCrLf
	sHTML = sHTML & "</tfoot>" & vbCrLf
	sHTML = sHTML & "</table>" & vbCrLf

	GetHtmlOrderList = sHTML
End Function

'******************************************************************************
'概　要：レポート（受注・登録者一覧）の登録者部分のＨＴＭＬ取得
'引　数：
'備　考：
'使　用：社内/report/order_staff_list.asp
'更　新：2007/09/20 LIS K.Kokubo 作成
'******************************************************************************
Function GetHtmlStaffList(ByRef rDB, ByVal vYM, ByVal vBranchCode)
	Dim sSQL
	Dim oRS
	Dim sError
	Dim flgQE

	Dim dbStaffName
	Dim dbJobTypeName
	Dim dbCompanyName
	Dim dbSelectionTypeName

	Dim sHTML
	Dim sOldOrderCode
	Dim sOldCompany
	Dim sOldStaff
	Dim sStyle
	Dim iRow
	Dim iAbsPos
	Dim iFltRec
	Dim iFltRec2

	sOldOrderCode = ""
	sOldCompany = ""
	sOldStaff = ""
	sSQL = "up_LstRptStaff '" & vYM & "', '" & vBranchCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	sHTML = sHTML & "<table class=""orderstafflist"" border=""0"" style=""float:left; width:380px;"">" & vbCrLf
	sHTML = sHTML & "<thead>" & vbCrLf
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<th colspan=""5"">登録者</th>"
	sHTML = sHTML & "</tr>" & vbCrLf
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td>No.</td>"
	sHTML = sHTML & "<td>氏名</td>"
	sHTML = sHTML & "<td>職種</td>"
	sHTML = sHTML & "<td>エントリー</td>"
	sHTML = sHTML & "<td>結果</td>"
	sHTML = sHTML & "</tr>" & vbCrLf
	sHTML = sHTML & "</thead>" & vbCrLf
	sHTML = sHTML & "<tbody>" & vbCrLf
	Do While GetRSState(oRS) = True
		sStyle = ""
		iFltRec = 0

		dbStaffName = oRS.Collect("StaffName")
		dbJobTypeName = Replace(oRS.Collect("JobTypeName"), vbCrLf, "<br>")
		dbCompanyName = oRS.Collect("CompanyName")
		dbSelectionTypeName = oRS.Collect("SelectionTypeName")

		'求職者毎の選考数を取得
		If sOldStaff <> dbStaffName Then
			iAbsPos = oRS.AbsolutePosition
			oRS.Filter = "StaffCode = '" & oRS.Collect("StaffCode") & "'"
			iFltRec = oRS.RecordCount
			oRS.Filter = 0
			oRS.AbsolutePosition = iAbsPos
		End If

		If sOldStaff <> dbStaffName And sOldStaff <> "" Then
			sStyle = "border-top:1px dotted #666666;"
		End If

		If sOldStaff = dbStaffName And sOldStaff <> "" Then
			dbStaffName = ""
			dbJobTypeName = ""
		End If

		If iFltRec > 1 Then
			sHTML = sHTML & "<td rowspan=""" & iFltRec & """ style=""border-left:1px solid #000000;" & sStyle & """>" & iFltRec & "</td>"
			sHTML = sHTML & "<td rowspan=""" & iFltRec & """ style=""" & sStyle & """>" & dbStaffName & "</td>"
			sHTML = sHTML & "<td rowspan=""" & iFltRec & """ style=""" & sStyle & """>" & dbJobTypeName & "</td>"
		ElseIf iFltRec = 1 Then
			sHTML = sHTML & "<td style=""border-left:1px solid #000000;" & sStyle & """>" & iFltRec & "</td>"
			sHTML = sHTML & "<td style=""" & sStyle & """>" & dbStaffName & "</td>"
			sHTML = sHTML & "<td style=""" & sStyle & """>" & dbJobTypeName & "</td>"
		End If
		sHTML = sHTML & "<td style=""" & sStyle & " border-bottom:1px dotted #666666;"">" & dbCompanyName & "</td>"
		sHTML = sHTML & "<td style=""" & sStyle & " border-bottom:1px dotted #666666; border-right:1px solid #333333;"">" & dbSelectionTypeName & "</td>"
		sHTML = sHTML & "</tr>" & vbCrLf

		sOldOrderCode = oRS.Collect("OrderCode")
		sOldStaff = oRS.Collect("StaffName")
		oRS.MoveNext
	Loop
	sHTML = sHTML & "<tr class=""setwidth"">" & vbCrLf
	sHTML = sHTML & "<th style=""width:30px;""></th>"
	sHTML = sHTML & "<th style=""width:80px;""></th>"
	sHTML = sHTML & "<th style=""width:80px;""></th>"
	sHTML = sHTML & "<th style=""width:80px;""></th>"
	sHTML = sHTML & "<th style=""width:30px;""></th>"
	sHTML = sHTML & "</tr>" & vbCrLf
	sHTML = sHTML & "</tbody>" & vbCrLf
	sHTML = sHTML & "<tfoot>" & vbCrLf
	sHTML = sHTML & "<tr><th colspan=""5""></th></tr>" & vbCrLf
	sHTML = sHTML & "</tfoot>" & vbCrLf
	sHTML = sHTML & "</table>" & vbCrLf

	GetHtmlStaffList = sHTML
End Function
%>
