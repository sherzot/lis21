<%
'*******************************************************************************
'�T�@�v�F��Ɩ����x�A���P�[�g�R���e���c�́u��Ƃ̃A���P�[�g��,���X�񓚁v������HTML���擾
'���@���FvSeq				�F�A���P�[�g��
'�@�@�@�FvAnswerDay			�F�A���P�[�g�񓚓�(���t�^)
'�@�@�@�FvSatisfactionPoint	�F�����x�|�C���g
'�@�@�@�FvOpinion			�F�ӌ�
'�@�@�@�FvLISAnswer			�F���X�񓚓��e
'�@�@�@�FvQuestionnaireID	�F��Ɩ����x�A���P�[�gID
'�߂�l�FString
'���@�l�F
'�X�@�V�F2009/06/01 LIS K.Kokubo �쐬
'*******************************************************************************
Function htmlSatisfaction(ByVal vSeq, ByVal vAnswerDay, ByVal vSatisfactionPoint, ByVal vOpinion, ByVal vLISAnswer, ByVal vQuestionnaireID)
	Dim sHTML

	sHTML = ""

	'<��Ɖ񓚕���>
	sHTML = sHTML & "<div style=""border-bottom:1px dashed #ccc; margin:0 0 5px;"">"
	'�^�C�g��
	sHTML = sHTML & "<div style=""border-bottom: 3px double #999999;"">"
	sHTML = sHTML & "<div style=""float:left;width:74%;background-color:transparent;"">"
	sHTML = sHTML & "<div style=""padding:4px;font-size:16px;"">"
	Select Case vSatisfactionPoint
		Case 1: sHTML = sHTML & "<span style=""font-weight:bold;"">�����x�F<span style=""color:gold;"">��</span></span>�@���Ȃ�s��"
		Case 2: sHTML = sHTML & "<span style=""font-weight:bold;"">�����x�F<span style=""color:gold;"">����</span></span>�@�ǂ��炩�Ƃ����ƕs��"
		Case 3: sHTML = sHTML & "<span style=""font-weight:bold;"">�����x�F<span style=""color:gold;"">������</span></span>�@�ӂ�"
		Case 4: sHTML = sHTML & "<span style=""font-weight:bold;"">�����x�F<span style=""color:gold;"">��������</span></span>�@�ǂ��炩�Ƃ����Ɩ���"
		Case 5: sHTML = sHTML & "<span style=""font-weight:bold;"">�����x�F<span style=""color:gold;"">����������</span></span>�@���Ȃ薞��"
	End Select
	sHTML = sHTML & "&nbsp;(��" & vSeq & "��)"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div style=""float:right;width:25%;background-color:transparent;text-align:right;""><div style=""padding:4px;"">�񓚓��F" & GetDateStr(vAnswerDay, "/") & "</div></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	sHTML = sHTML & "</div>"
	'���e
	sHTML = sHTML & "<div style=""padding:4px;""><b>L</b><p class=""m0"" style=""float:right;width:705px;"">" & Replace(vOpinion, vbCrLf, "<br>") & "</p></div>"
	sHTML = sHTML & "<div clear=""both""></div></div>"
	'</��Ɖ񓚕���>

	sHTML = sHTML & "<div style=""margin:0 0 5px 15px; border-bottom:1px dashed #ccc;"">"

	'<���X�񓚕���>
	If vLISAnswer <> "" Then
		'�^�C�g��

		sHTML = sHTML & "<div style=""padding:4px;font-size:16px;font-weight: bold;"">�����ƃi�r����̉�</div>"

		'���e
		sHTML = sHTML & "<div style=""padding:4px;""><b>L</b><p class=""m0"" style=""float:right;width:690px;"">" & Replace(vLISAnswer, vbCrLf, "<br>") & "</p></div>"
		sHTML = sHTML & "<div clear=""both""></div>"
	End If
	'</���X�񓚕���>

	sHTML = sHTML & "</div>"

	htmlSatisfaction = sHTML
End Function
%>
