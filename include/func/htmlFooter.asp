<%
'*******************************************************************************
'�T�@�v�FHTML������DOCTYPE�`body�^�O�܂ł��擾
'���@���FvHTML	�F</body>�̒��O�ɑ}������HTML
'�o�@�́F
'�߂�l�FString
'���@�l�F
'���@���F2010/05/11 LIS K.Kokubo �쐬
'*******************************************************************************
Function htmlFooter(ByVal vHTML)
	Dim sHTML

	sHTML = vbCrLf & vHTML & vbCrLf & "</body></html>"

	htmlFooter = sHTML
End Function
%>
