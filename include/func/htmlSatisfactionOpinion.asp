<%
'*******************************************************************************
'�T�@�v�F��Ɩ����x�A���P�[�g�R���e���c�́u�F����̈ӌ��v������HTML���擾
'���@���FvQuestionnaireID	�F��Ɩ����x�A���P�[�gID
'�߂�l�FString
'���@�l�F
'�X�@�V�F2009/06/01 LIS K.Kokubo �쐬
'*******************************************************************************
Function htmlSatisfactionOpinion(ByVal vQuestionnaireID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	'DB
	Dim dbPenName
	Dim dbOpinion
	Dim dbRegistDay

	Dim sHTML

	sHTML = ""

	sSQL = ""
	sSQL = sSQL & "/* �����ƃi�r ��Ɩ����x�A���P�[�g �R�����g�ꗗ�擾 */" & vbCrLf
	sSQL = sSQL & "EXEC up_LstCMPQuestionnaireOpinion '" & vQuestionnaireID & "','1','1';"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		sHTML = sHTML & "<div style=""margin:0 0 5px 15px; border-bottom:1px dashed #ccc;"">"
		sHTML = sHTML & "<div style=""padding:4px;font-size:16px; font-weight:bold;color:orange;"">�F���񂩂�̈ӌ�</div>"

		Do While GetRSState(oRS) = True
			dbPenName = ChkStr(oRS.Collect("PenName"))
			dbOpinion = ChkStr(oRS.Collect("Opinion"))
			dbRegistDay = oRS.Collect("RegistDay")

			'�R�����g
			sHTML = sHTML & "<div style=""padding:4px;"">"
			sHTML = sHTML & "<p class=""m0"">�y���l�[���F" & dbPenName & "</p>"
			sHTML = sHTML & "<div style=""margin-left:15px;"">"
			sHTML = sHTML & "<b>L</b><p class=""m0"" style=""float:right;width:675px;"">" & Replace(Trim(dbOpinion) & "(" & GetDateStr(dbRegistDay, "/") & ")", vbCrLf, "<br>") & "</p>"
			sHTML = sHTML & "</div>"
			sHTML = sHTML & "<div clear=""both""></div></div>"


			oRS.MoveNext
		Loop
		sHTML = sHTML & "</div>"
	End If
	Call RSClose(oRS)

	htmlSatisfactionOpinion = sHTML
End Function
%>
