<%
'*******************************************************************************
'�T�@�v�F���E�҂̉���o�^��ʂ̃A���P�[�gHTML���擾
'���@���F
'�o�@�́F
'�߂�l�FString
'���@�l�F
'���@���F2011/07/06 LIS K.Kokubo �쐬
'*******************************************************************************
Function htmlStaffRegQuestionnaire()
	Dim sHTML

	Dim frmQ1
	Dim aQ1(6),idx,tmpAry

	frmQ1 = GetForm("frmq1",2)

	'<�`�F�b�N�{�b�N�X or ���a�I�{�^���̃f�t�H���g�ݒ�>
	tmpAry = Split(Replace(frmQ1," ",""),",")
	For idx = 0 To UBound(tmpAry)
		aQ1(tmpAry(idx)) = " checked"
	Next
	'</�`�F�b�N�{�b�N�X or ���a�I�{�^���̃f�t�H���g�ݒ�>


	sHTML = sHTML & "<table class=""pattern1_1"" >"
	sHTML = sHTML & "<tbody>"

	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<th id=""thenq"" class=""first_th"">�A���P�[�g�ɂ����͂�������<br>�i���o�^�̂��������́H�j</th>"
	sHTML = sHTML & "<td class=""first_td"">"
	sHTML = sHTML & "<ul class=""left"" style=""margin-right:80px;"">"
	sHTML = sHTML & "<li><label><input name=""frmq1"" type=""checkbox"" value=""0""" & aQ1(0) & " onclick=""blur();if(needcount)needcount();""> ��]���鋁�l�ւ̉���</label></li>"
	sHTML = sHTML & "<li><label><input name=""frmq1"" type=""checkbox"" value=""1""" & aQ1(1) & " onclick=""blur();if(needcount)needcount();""> �]�E�T�|�[�g(�l�ޏЉ�E�h��)����]</label></li>"
	sHTML = sHTML & "<li><label><input name=""frmq1"" type=""checkbox"" value=""2""" & aQ1(2) & " onclick=""blur();if(needcount)needcount();""> �������E�E���o�����쐬�̗��p</label></li>"
	sHTML = sHTML & "<li><label><input name=""frmq1"" type=""checkbox"" value=""3""" & aQ1(3) & " onclick=""blur();if(needcount)needcount();""> �������̃R���e���c���p</label></li>"
	sHTML = sHTML & "</ul>"
	sHTML = sHTML & "<ul>"
	sHTML = sHTML & "<li><label><input name=""frmq1"" type=""checkbox"" value=""4""" & aQ1(4) & " onclick=""blur();if(needcount)needcount();""> �g�ѓd�b�E�X�}�[�g�t�H���œ]�E����</label></li>"
	sHTML = sHTML & "<li><label><input name=""frmq1"" type=""checkbox"" value=""5""" & aQ1(5) & " onclick=""blur();if(needcount)needcount();""> �R���r�j����@�\����]</label></li>"
	sHTML = sHTML & "<li><label><input name=""frmq1"" type=""checkbox"" value=""6""" & aQ1(6) & " onclick=""blur();if(needcount)needcount();""> ���̑�</label> �i <input id=""txtQ1"" name=""frmq1other"" type=""text"" maxlength=""200"" value="""" style=""width:180px;"" onkeyup=""if(needcount)needcount();""> �j</li>"
	sHTML = sHTML & "</ul>"
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"

	sHTML = sHTML & "</tbody>"
	sHTML = sHTML & "</table>"

	htmlStaffRegQuestionnaire = sHTML
End Function
%>
