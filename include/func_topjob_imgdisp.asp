<%
Function DspTopJobImage()
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbApplicationCode
	Dim dbSeq
	Dim dbOrderCode
	Dim dbCompanyCode
	Dim dbCompanyName
	Dim dbPicType	'写真種類 [1]企業写真 [2]オプションで登録した写真

	Dim idx
	Dim sImgSrc

	sSQL = sSQL & "SELECT DISTINCT VWNLV.UserCode AS CompanyCode "
	sSQL = sSQL & "FROM vw_NaviLicense_Valid AS VWNLV "
	sSQL = sSQL & "INNER JOIN LicenseData_Option AS LDO "
	sSQL = sSQL & "ON VWNLV.ApplicationCode = LDO.ApplicationCode "
	sSQL = sSQL & "AND LDO.OptionCode = 2 "
	sSQL = sSQL & "INNER JOIN OPTNaviTopPicture AS ONTP "
	sSQL = sSQL & "ON ONTP.ApplicationCode = LDO.ApplicationCode "
	sSQL = sSQL & "AND ONTP.Seq = LDO.Seq "
	sSQL = sSQL & "AND (CONVERT(VARCHAR(8), GETDATE(), 112) BETWEEN ONTP.StartDay AND ONTP.EndDay) "
	sSQL = sSQL & "INNER JOIN vw_OrderCode AS VWOC "
	sSQL = sSQL & "ON VWNLV.UserCode = VWOC.CompanyCode "
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	'トップ写真掲載オプションが無い場合は出力しない
	If GetRSState(oRS) = False Then Exit Function

	sSQL = "EXEC up_LstOPTNaviTopPicture_Dsp"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	idx = 1

	Do While idx <= 4
		sImgSrc = ""
		dbApplicationCode = ""
		dbSeq = ""
		dbOrderCode = ""
		dbCompanyCode = ""
		dbCompanyName = ""
		dbPicType = ""

		If GetRSState(oRS) Then
			dbApplicationCode = oRS.Collect("ApplicationCode")
			dbSeq = oRS.Collect("Seq")
			dbOrderCode = oRS.Collect("OrderCode")
			dbCompanyCode = oRS.Collect("CompanyCode")
			dbCompanyName = oRS.Collect("CompanyName")
			dbPicType = oRS.Collect("PicType")

			If dbPicType = "2" Then
				sImgSrc = "/company/imgdsp_navitop.asp?applicationcode=" & dbApplicationCode & "&amp;seq=" & dbSeq
			Else
				sImgSrc = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=1"
			End If

			Response.Write "<div style=""width:150px; float:left; margin:0px;"">"
			Response.Write "<a href=""" & HTTP_CURRENTURL & "order/company_order.asp?poc=" & dbOrderCode & """><img src=""" & sImgSrc & """ alt=""" & dbCompanyName & """ width=""145"" height=""108"" border=""1"" style=""border:1px solid #666666;""></a><br>"
			Response.Write "<div style=""padding:2px 4px 0px 0px;text-align:left;font-size:12px;"">"
			Response.Write "<a href=""" & HTTP_CURRENTURL & "order/company_order.asp?poc=" & dbOrderCode & """ class=""topdecnone"" style=""font-size:11px;"">" & dbCompanyName & "</a><br>"
			Response.Write "</div>"
			Response.Write "</div>"
		End If

		idx = idx + 1
		If GetRSState(oRS) = True Then oRS.MoveNext
	Loop

	Response.Write "<div style=""clear:both;""></div>"

'	Do While idx <= 8
'		sImgSrc = ""
'		dbApplicationCode = ""
'		dbSeq = ""
'		dbOrderCode = ""
'		dbCompanyCode = ""
'		dbCompanyName = ""
'		dbPicType = ""
'
'		If GetRSState(oRS) = True Then
'			dbApplicationCode = oRS.Collect("ApplicationCode")
'			dbSeq = oRS.Collect("Seq")
'			dbOrderCode = oRS.Collect("OrderCode")
'			dbCompanyCode = oRS.Collect("CompanyCode")
'			dbCompanyName = oRS.Collect("CompanyName")
'			dbPicType = oRS.Collect("PicType")
'
'			If dbPicType = "2" Then
'				sImgSrc = "/company/imgdsp_navitop.asp?applicationcode=" & dbApplicationCode & "&amp;seq=" & dbSeq
'			Else
'				sImgSrc = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=1"
'			End If
'
'			Response.Write "<div style=""width:150px; float:left; margin:0px;"">"
'			Response.Write "<a href=""" & HTTP_CURRENTURL & "order/company_order.asp?poc=" & dbOrderCode & """><img src=""" & sImgSrc & """ alt=""" & dbCompanyName & """ width=""145"" height=""108"" border=""1"" style=""border:1px solid #666666;""></a><br>"
'			Response.Write "<div style=""padding:2px 4px 0px 0px;text-align:left;font-size:12px;"">"
'			Response.Write "<a href=""" & HTTP_CURRENTURL & "order/company_order.asp?poc=" & dbOrderCode & """ class=""topdecnone"" style=""font-size:11px;"">" & dbCompanyName & "</a><br>"
'			Response.Write "</div>"
'			Response.Write "</div>"
'		End If
'
'		idx = idx + 1
'		If GetRSState(oRS) = True Then oRS.MoveNext
'	Loop

	Call RSClose(oRS)
End Function

Function DspTopJobImage2()
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbApplicationCode
	Dim dbSeq
	Dim dbOrderCode
	Dim dbCompanyCode
	Dim dbCompanyName

	Dim idx
	Dim sImgSrc

	sSQL = "EXEC up_LstOPTNaviTopPicture_Dsp"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	idx = 1

	Do While idx <= 4
		sImgSrc = ""
		dbApplicationCode = ""
		dbSeq = ""
		dbOrderCode = ""
		dbCompanyCode = ""
		dbCompanyName = ""

		If GetRSState(oRS) Then
			dbApplicationCode = oRS.Collect("ApplicationCode")
			dbSeq = oRS.Collect("Seq")
			dbOrderCode = oRS.Collect("OrderCode")
			dbCompanyCode = oRS.Collect("CompanyCode")
			dbCompanyName = oRS.Collect("CompanyName")

			If ChkStr(dbSeq) <> "" Then
				sImgSrc = "/company/imgdsp_navitop.asp?applicationcode=" & dbApplicationCode & "&amp;seq=" & dbSeq
			Else
				sImgSrc = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=1"
			End If
%>
		<div style="width:150px; float:left; margin:0px;">
			<a href="<%= HTTP_CURRENTURL %>order/company_order.asp?poc=<%= dbOrderCode %>"><img src="<%= sImgSrc %>" alt="<%= dbCompanyName %>" width="145" height="108" border="1" style="border:1px solid #666666;"></a><br>
		</div>
<%
		End If

		idx = idx + 1
		If GetRSState(oRS) = True Then oRS.MoveNext
	Loop
%>
		<div style="clear:both;"></div>

<%
	Call RSClose(oRS)
End Function
%>
