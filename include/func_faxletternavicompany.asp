<%
'*****************************************************
'** ��ƌ��������ƃi�r������̃t�@�b�N�X�ʒm�@�\
'** 
'** �ϐ��ꗗ
'**		vCn	:	��Ж�
'**		vPn	:	��ƒS���Ҏ���
'**		vSc	:	�X�^�b�t�R�[�h
'**		vOc	:	���l���R�[�h
'**		vSj	:	����
'**		vBd	:	�{��
'**		vFn	:	�t�@�b�N�X�ԍ�
'**	�Ԃ�l
'** 	����I�� : true �A�ُ� : Err.Number
'*****************************************************
function FaxLetterNaviCompany(vCn,vPn,vSc,vOc,vSj,vBd,vFn)
	
	On Error Resume Next
	
'*******************************************
'FAX���M�p������Excel�t�H�[�}�b�g�Ő�������B
'*******************************************
'	ReportFolder			 �e�풠�[�t�H�[�}�b�g�ۊǏꏊ personnel.asp�ɋL�q
	'�쐬�t�@�C����
	Dim wOutFileName	:	wOutFileName	=	"����" & vOc & year(Now()) & Month(Now()) & Day(Now()) & Hour(Now()) & Minute(Now()) & Second(Now())
	'�ۑ���̃t�H���_
	Dim wSaveFolder		:	wSaveFolder		=	"\\192.168.10.61\fax���M����"
	
	
	response.write Err.Number
	response.write Err.Description
	
	Dim wErrNo
	Dim wMsg
	Dim Xlsx1
	
	'ExcelCreator�I�u�W�F�N�g����
	Set Xlsx1 = Server.CreateObject("XlsxCrt.XlsxCrtCtrl.1")
	
	'Excel�t�@�C���i����`�[�j�I�[�o���C�I�[�v��
	Xlsx1.OpenBook wSaveFolder & "\" & wOutFileName & ".xlsx", ReportFolder & "\�����ƃi�r����A���ʒm�����i��ƌ����j.xlsx"
	
	Xlsx1.SheetNo = 0
    
    Xlsx1.Cell("A6").Value = vCn & "�l�B"
    Xlsx1.Cell("A7").Value = "���E�ҁi" & vSc & ")����A�����ƃi�r��ʂ���" & _
    						vOc & "�̋��l�ɂ��ĘA��������܂����̂łe�`�w�ɂĂ��`���v���܂��B"
    
    Xlsx1.Cell("B11").Value = vSj
    Xlsx1.Cell("B12").Value = vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd
    
	If wErrNo <> 0 Then
		FaxLetterNaviCompany = "ExcelCreator3.6 �G���[���b�Z�[�W�F" & Xlsx1.ErrorMessage
		exit function
	End If
	Xlsx1.CloseBook
	Set Xlsx1 = Nothing
	
'*******************************************
'FAX���M�pCSV�t�@�C���𐶐�����B
'*******************************************
	Dim wCsvFolder	:	wCsvFolder = "\\192.168.10.61\CsvShare"
	
	Set objFS = CreateObject("Scripting.FileSystemObject")
	Set objFolder = objFS.GetFolder(wCsvFolder)
	Set objFile = objFolder.CreateTextFile(wOutFileName & ".csv")
	objFile.WriteLine("""C:\FAX���M����\" & wOutFileName & ".xlsx" & """,""" & vFn & """,""" & vPn & ""","" " & "�l" & " "",""" & vCn & ""","""","""",""""")
	objFile.Close
	
'*******************************************
'�G���[�`�F�b�N
'*******************************************
	response.write wSaveFolder & "\" & wOutFileName
	response.write "<br>"
	response.write Err.Number
	response.write Err.Description
	
end function
%>