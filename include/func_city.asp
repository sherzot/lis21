<%
'******************************************************************************
'�T�@�v�F�n����݃y�[�W�̌��������\���֐�
'���@���FvCityURL			�F�n��y�[�W�̂t�q�k�@[��]minatoku.asp
'�@�@�@�FvCity				�F�n�於�@[��]�`��
'�@�@�@�FvChikuCode			�F�n��R�[�h�@[��]akasaka��B_Station.Chiku
'�@�@�@�FvWorkingTypeCode	�F�ٗp�`�ԃR�[�h�@[��]002
'�@�@�@�FvJobTypeCode		�F�E��R�[�h�@[��]01
'�@�@�@�FvChikuSearchFlag	�F�n�挟���\���ہ@[0]��\�� [1]�\��
'�߂�l�F
'�쐬���F2006/12/13
'�쐬�ҁFLis Kokubo
'���@�l�F
'�X�@�V�F
'******************************************************************************
Sub DspCitySearchConditionHtml(ByVal vCityURL, ByVal vPrefectureCode, ByVal vCity, ByVal vChikuCode, ByVal vWorkingTypeCode, ByVal vJobTypeCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim idx

	Dim sWTName
	Dim sJTName
	Dim sChikuName
	Dim sCondition1:	sCondition1 = ""
	Dim sCondition2:	sCondition2 = ""

	sWTName = GetDetail("WorkingType", vWorkingTypeCode)
	sJTName = GetJobTypeBig(Left(vJobTypeCode, 2))
	sChikuName = GetChiku(vPrefectureCode, vChikuCode)
	If sWTName <> "" Then
		If sCondition1 <> "" Then sCondition1 = sCondition1 & "�A"
		If sCondition2 <> "" Then sCondition2 = sCondition2 & "�A"
		sCondition2 = sCondition2 & sWTName
	End If
	If sJTName <> "" Then
		If sCondition1 <> "" Then sCondition1 = sCondition1 & "�A"
		If sCondition2 <> "" Then sCondition2 = sCondition2 & "�A"
		sCondition1 = sCondition1 & sJTName
	End If
	If sChikuName <> "" Then
		If sCondition1 <> "" Then sCondition1 = sCondition1 & "�A"
		If sCondition2 <> "" Then sCondition2 = sCondition2 & "�A"
		sCondition1 = sCondition1 & sChikuName
		sCondition2 = sCondition2 & sChikuName
	End If
	If sCondition1 <> "" Then sCondition1 = "&nbsp;&nbsp;<span style=""font-weight:normal;"">()���́A�����u<span style=""font-weight:bold;"">" & sCondition1 & "</span>�v�ł̋��l�[�����B</span>"
	If sCondition2 <> "" Then sCondition2 = "&nbsp;&nbsp;<span style=""font-weight:normal;"">()���́A�����u<span style=""font-weight:bold;"">" & sCondition2 & "</span>�v�ł̋��l�[�����B</span>"

%>
<a name="#search"></a>
<h2 class="ssubtitle">���������I��</h2>
<div class="subcontent">
	<table class="citysearch" border="0" cellspacing="0">
		<tbody>
<%
	'*******************************************************************************
	'�E��啪�� start
	'*******************************************************************************
%>
		<tr>
			<th class="citysearch" valign="top">
				<p style="float:left; width:60px; margin:0px;">�E��</p>
				<p style="float:left; margin:0px;"><%= sCondition2 %></p>
				<p class="citynosearch"><a href="<%= BASEURL %>city/<%= vCityURL %>?wt=<%= vWorkingTypeCode %>&amp;chiku=<%= vChikuCode %>&amp;jt=" title="<%= vCity %>�̓]�E���l���" style="font-weight:normal;">�w�肵�Ȃ�</a></p>
				<br clear="all">
			</th>
		</tr>
		<tr>
			<td class="citysearch">
<%
	sSQL = "up_GetListCityJobTypeBig '" & vPrefectureCode & "', '" & vCity & "', '" & vChikuCode & "', '" & vWorkingTypeCode & "', '" & vJobTypeCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	idx = 1
	Do While GetRSState(oRS) = True
		If oRS.Collect("JobTypeCode") = vJobTypeCode Then
			Response.Write "<p class=""citysearch"" style=""color:#ff0000;"">" & oRS.Collect("JobTypeName") & "(" & oRS.Collect("Cnt") & ")</p>"
		ElseIf oRS.Collect("Cnt") > 0 Then
			Response.Write "<p class=""citysearch""><a href=""" & BASEURL & "city/" & vCityURL & "?wt=" & vWorkingTypeCode & "&amp;chiku=" & vChikuCode & "&amp;jt=" & oRS.Collect("JobTypeCode") & """ title=""" & vCity & "��" & oRS.Collect("JobTypeName") & "�̓]�E���l���"" style=""" & GetConditionStr(oRS.Collect("JobTypeCode"), vJobTypeCode, "color:#ff0000;") & """>" & oRS.Collect("JobTypeName") & "(" & oRS.Collect("Cnt") & ")</a></p>"
		Else
			Response.Write "<p class=""citysearch"">" & oRS.Collect("JobTypeName") & "</p>"
		End If
		If idx Mod 4 = 0 Then Response.Write "<br clear=""all"">"
		oRS.MoveNext
		If GetRSState(oRS) = False And idx Mod 4 <> 0 Then Response.Write "<br clear=""all"">"
		idx = idx + 1
	Loop
	Call RSClose(oRS)
%>
			</td>
		</tr>
<%
	'*******************************************************************************
	'�E��啪�� end
	'*******************************************************************************

	'*******************************************************************************
	'�ٗp�`�� start
	'*******************************************************************************
%>
		<tr>
			<th class="citysearch" valign="top">
				<p style="float:left; width:60px; margin:0px;">�ٗp�`��</p>
				<p style="float:left; margin:0px;"><%= sCondition1 %></p>
				<p class="citynosearch"><a href="<%= BASEURL %>city/<%= vCityURL %>?wt=&amp;chiku=<%= vChikuCode %>&amp;jt=<%= vJobTypeCode %>" title="<%= vCity %>�̓]�E���l���" style="font-weight:normal;">�w�肵�Ȃ�</a></p>
				<br clear="all">
			</th>
		</tr>
		<tr>
			<td class="citysearch">
<%
	sSQL = "up_GetListCityWorkingType '" & vPrefectureCode & "', '" & vCity & "', '" & vChikuCode & "', '" & vWorkingTypeCode & "', '" & vJobTypeCode & "'"
	Set oRS = QEXE(dbconn, sSQL)
	idx = 1
	Do While GetRSState(oRS) = True
		If oRS.Collect("WorkingTypeCode") = vWorkingTypeCode Then
			Response.Write "<p class=""citysearch"" style=""color:#ff0000;"">" & oRS.Collect("WorkingTypeName") & "(" & oRS.Collect("Cnt") & ")</p>"
		ElseIf oRS.Collect("Cnt") > 0 Then
			Response.Write "<p class=""citysearch""><a href=""" & BASEURL & "city/" & vCityURL & "?wt=" & oRS.Collect("WorkingTypeCode") & "&amp;chiku=" & vChikuCode & "&amp;jt=" & vJobTypeCode & """ title=""" & vCity & "�̓]�E���l���"" style=""" & GetConditionStr(oRS.Collect("WorkingTypeCode"), vWorkingTypeCode, "color:#ff0000;") & """>" & oRS.Collect("WorkingTypeName") & "(" & oRS.Collect("Cnt") & ")</a></p>"
		Else
			Response.Write "<p class=""citysearch"">" & oRS.Collect("WorkingTypeName") & "</p>"
		End If
		If idx Mod 4 = 0 Then Response.Write "<br clear=""all"">"
		oRS.MoveNext
		If GetRSState(oRS) = False And idx Mod 4 <> 0 Then Response.Write "<br clear=""all"">"
		idx = idx + 1
	Loop
	Call RSClose(oRS)
%>
			</td>
		</tr>
<%
	'*******************************************************************************
	'�ٗp�`�� end
	'*******************************************************************************

	'*******************************************************************************
	'�n�� start
	'*******************************************************************************
	If vCity = "�`��" Then
%>
		<tr>
			<th class="citysearch" valign="top">
				<p class="citysearch">�G���A</p>
				<p class="citynosearch"><a href="<%= BASEURL %>city/<%= vCityURL %>?wt=<%= vWorkingTypeCode %>&amp;chiku=&amp;jt=<%= vJobTypeCode %>" title="<%= vCity %>�̓]�E���l���" style="font-weight:normal;">�w�肵�Ȃ�</a></p>
				<br clear="all">
			</th>
		</tr>
		<tr>
			<td class="citysearch">
<%
		sSQL = "up_GetListCityChiku '" & vPrefectureCode & "', '" & vCity & "', '" & vChikuCode & "', '" & vWorkingTypeCode & "', '" & vJobTypeCode & "'"
		Set oRS = QEXE(dbconn, sSQL)
		idx = 1
		Do While GetRSState(oRS) = True
			If oRS.Collect("ChikuCode") = vChikuCode Then
				Response.Write "<p class=""citysearch"" style=""color:#ff0000;"">" & oRS.Collect("ChikuName") & "(" & oRS.Collect("Cnt") & ")</p>"
			ElseIf oRS.Collect("Cnt") > 0 Then
				Response.Write "<p class=""citysearch""><a href=""" & BASEURL & "city/" & vCityURL & "?wt=" & vWorkingTypeCode & "&amp;chiku=" & oRS.Collect("ChikuCode") & "&amp;jt=" & vJobTypeCode & """ title=""" & vCity & "�̓]�E���l���"" style=""" & GetConditionStr(oRS.Collect("ChikuCode"), vChikuCode, "color:#ff0000;") & """>" & oRS.Collect("ChikuName") & "(" & oRS.Collect("Cnt") & ")</a></p>"
			Else
				Response.Write "<p class=""citysearch"">" & oRS.Collect("ChikuName") & "</p>"
			End If
			If idx Mod 4 = 0 Then Response.Write "<br clear=""all"">"
			oRS.MoveNext
			If GetRSState(oRS) = False And idx Mod 4 <> 0 Then Response.Write "<br clear=""all"">"
			idx = idx + 1
		Loop
		Call RSClose(oRS)
	End If
	'*******************************************************************************
	'�n�� end
	'*******************************************************************************
%>
		</tbody>
	</table>
</div>
<%
End Sub

'******************************************************************************
'�T�@�v�FvComp1��vComp2���������ꍇ�́AvStr���擾
'���@���FvComp1	�F���r
'�@�@�@�FvComp2	�F��r
'�@�@�@�FvStr	�FvComp1��vComp2�����������ɕԂ�������
'�߂�l�F
'�쐬���F2006/12/13
'�쐬�ҁFLis Kokubo
'���@�l�F
'�X�@�V�F
'******************************************************************************
Function GetConditionStr(ByVal vComp1, ByVal vComp2, ByVal vStr)
	GetConditionStr = ""
	If vComp1 = vComp2 Then GetConditionStr = vStr
End Function

'******************************************************************************
'�T�@�v�F�n����݃y�[�W�̋��l�[�ꗗ�\���֐�
'���@���FvCityURL			�F�n��y�[�W�̂t�q�k�@[��]minatoku.asp
'�@�@�@�FvPrefectureCode	�F�s���{���R�[�h�@[��]013
'�@�@�@�FvCity				�F�n�於�@[��]�`��
'�@�@�@�FvChikuCode			�F�n��R�[�h�@[��]akasaka��B_Station.Chiku
'�@�@�@�FvWorkingTypeCode	�F�ٗp�`�ԃR�[�h�@[��]002
'�@�@�@�FvJobTypeCode		�F�E��R�[�h�@[��]01
'�@�@�@�FvPage				�F�y�[�W
'�߂�l�F
'�쐬���F2006/12/13
'�쐬�ҁFLis Kokubo
'���@�l�F
'�X�@�V�F
'******************************************************************************
Sub DspCityOrderListHtml(ByVal vCityURL, ByVal vPrefectureCode, ByVal vCity, ByVal vChikuCode, ByVal vWorkingTypeCode, ByVal vJobTypeCode, ByVal vPage)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sPageCtrl

	Response.Write "<h2 class=""ssubtitle"">" & vCity& "�̓]�E�E���l���</h2>"
	Response.Write "<div class=""subcontent"">"

	sSQL = "/* �����s23��l�C�G���A�̋��l���� */"
	sSQL = sSQL & "EXEC up_SearchOrderCity '" & vPrefectureCode & "', '" & vCity & "', '" & vChikuCode & "', '" & vWorkingTypeCode & "', '" & vJobTypeCode & "';"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

	If GetRSState(oRS) = True Then
		sPageCtrl = GetHtmlPageControlParam(dbconn, oRS, 3, vPage, HTTP_CURRENTURL & "city/" & vCityURL & "?wt=" & vWorkingTypeCode & "&chiku=" & vChikuCode & "&jt=" & vJobTypeCode, "")
		Response.Write sPageCtrl
		Response.Write "<div class=""line1"" style=""padding-bottom:5px;""></div>"
		Call DspOrderListDetail3(dbconn, oRS, 3, vPage, "")
		Response.Write "<div class=""line1""></div>"
		Response.Write sPageCtrl
	Else
		Response.Write "<p>���T���̋��l���E���d�����͂���܂���B</p>"
	End If

	Response.Write "</div>"
End Sub
%>
