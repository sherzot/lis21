<%
'****************************************************
'*	�֐����FFormatDate(�ȈՔ�)						*
'*	�����@�F���t�̌`����ϊ�����					*
'*	�����@�FYMD  �`�����t							*
'*	�@�@�@�@MODE 0:0�Â߂��s��Ȃ�					*
'*	�@�@�@�@     1:0�Â߂��s��						*
'*	�@�@�@�@SEP  �Z�p���[�^							*
'*	�߂�l�F�`�����ϊ����ꂽ���t					*
'****************************************************
function FormatDate( YMD, MODE, SEP )
	Dim wYYYY
	Dim wMM
	Dim wDD
	
	FormatDate = ""
	
	if Not isDate(YMD) then
		exit function
	end if
	
	if SEP = "" then
		SEP = "/"
	end if
	
	wYYYY = Year(YMD)
	if MODE = 1 then
		wMM = Right("00" & Month(YMD),2)
		wDD = Right("00" & Day(YMD),2)
	else
		wMM = Month(YMD)
		wDD = Day(YMD)
	end if
	
	FormatDate = wYYYY & SEP & wMM & SEP & wDD
end function
%>